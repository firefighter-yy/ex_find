"""Session-scoped SQLite index and streaming Excel preparation service.

The index is deliberately independent from both Qt and Excel COM.  The
preparation service consumes :class:`~ex_transform.excel_ingestion.WorksheetChunk`
objects and commits one source workbook at a time, so an interrupted or failed
workbook can never become searchable.
"""

from __future__ import annotations

import base64
import json
import os
import shutil
import sqlite3
import tempfile
import threading
import time
from collections.abc import Callable, Iterable
from concurrent.futures import Future
from dataclasses import dataclass, replace
from datetime import date, datetime
from datetime import time as datetime_time
from enum import Enum
from pathlib import Path
from types import TracebackType
from typing import Any

from .excel_ingestion import (
    ExcelReader,
    IngestionCancelledError,
    IngestionProgress,
    WorksheetChunk,
)


class IndexErrorBase(RuntimeError):
    """Base error for index lifecycle failures."""


class IndexTransactionError(IndexErrorBase):
    """Raised when a chunk is written outside an active file transaction."""


class FileStatus(str, Enum):
    """Persisted preparation states; string values keep SQL/API interop simple."""

    PREPARING = "preparing"
    READY = "ready"
    FAILED = "failed"
    CANCELLED = "cancelled"
    STALE = "stale"


@dataclass(frozen=True, slots=True)
class IndexedFile:
    id: int
    path: Path
    size: int | None
    mtime_ns: int | None
    status: str
    file_order: int
    error: str | None = None


@dataclass(frozen=True, slots=True)
class IndexedWorksheet:
    id: int
    file_id: int
    name: str
    sheet_index: int
    visible: bool
    first_row: int | None
    last_row: int | None
    first_column: int | None
    last_column: int | None


@dataclass(frozen=True, slots=True)
class IndexedCell:
    id: int
    file_id: int
    worksheet_id: int
    row: int
    column: int
    value: Any
    value_text: str | None
    is_formula: bool
    formula: str | None


@dataclass(frozen=True, slots=True)
class PreparationResult:
    """Summary returned after all requested files have been processed."""

    files: tuple[IndexedFile, ...]
    cancelled: bool = False

    @property
    def succeeded(self) -> tuple[IndexedFile, ...]:
        return tuple(item for item in self.files if item.status == "ready")

    @property
    def failed(self) -> tuple[IndexedFile, ...]:
        return tuple(item for item in self.files if item.status in {"failed", "stale"})


SCHEMA = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS session_metadata (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    path TEXT NOT NULL UNIQUE,
    size INTEGER,
    mtime_ns INTEGER,
    status TEXT NOT NULL CHECK (status IN ('preparing', 'ready', 'failed', 'cancelled', 'stale')),
    file_order INTEGER NOT NULL,
    error TEXT,
    created_at REAL NOT NULL,
    updated_at REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS worksheets (
    id INTEGER PRIMARY KEY,
    file_id INTEGER NOT NULL REFERENCES files(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    sheet_index INTEGER NOT NULL,
    visible INTEGER NOT NULL,
    first_row INTEGER,
    last_row INTEGER,
    first_column INTEGER,
    last_column INTEGER,
    UNIQUE(file_id, sheet_index)
);
CREATE TABLE IF NOT EXISTS physical_rows (
    id INTEGER PRIMARY KEY,
    worksheet_id INTEGER NOT NULL REFERENCES worksheets(id) ON DELETE CASCADE,
    row_number INTEGER NOT NULL,
    UNIQUE(worksheet_id, row_number)
);
CREATE TABLE IF NOT EXISTS cells (
    id INTEGER PRIMARY KEY,
    file_id INTEGER NOT NULL REFERENCES files(id) ON DELETE CASCADE,
    worksheet_id INTEGER NOT NULL REFERENCES worksheets(id) ON DELETE CASCADE,
    row_id INTEGER NOT NULL REFERENCES physical_rows(id) ON DELETE CASCADE,
    row_number INTEGER NOT NULL,
    column_number INTEGER NOT NULL,
    value TEXT,
    value_json TEXT,
    value_type TEXT NOT NULL,
    normalized_value TEXT,
    is_formula INTEGER NOT NULL DEFAULT 0,
    formula TEXT,
    UNIQUE(worksheet_id, row_number, column_number)
);
CREATE INDEX IF NOT EXISTS idx_files_status_order ON files(status, file_order);
CREATE INDEX IF NOT EXISTS idx_worksheets_file_order ON worksheets(file_id, sheet_index);
CREATE INDEX IF NOT EXISTS idx_rows_worksheet_number ON physical_rows(worksheet_id, row_number);
CREATE INDEX IF NOT EXISTS idx_cells_exact ON cells(normalized_value);
CREATE INDEX IF NOT EXISTS idx_cells_row ON cells(worksheet_id, row_number, column_number);
"""


def cleanup_stale_sessions(root: str | Path, *, older_than_seconds: float = 0) -> tuple[Path, ...]:
    """Remove abandoned ``excel-search-*`` session directories.

    A directory with a live ``.active`` marker is left alone.  The marker is
    removed by :meth:`SessionIndex.close`; old markers are considered stale.
    """
    parent = Path(root)
    if not parent.exists():
        return ()
    now = time.time()
    removed: list[Path] = []
    for directory in parent.glob("excel-search-*"):
        if not directory.is_dir():
            continue
        try:
            age = now - directory.stat().st_mtime
            marker = directory / ".active"
            if age < older_than_seconds:
                continue
            if marker.exists():
                try:
                    pid = int(marker.read_text(encoding="ascii").strip())
                    if pid == os.getpid():
                        continue
                    # os.kill(pid, 0) is available on Windows and only probes liveness.
                    os.kill(pid, 0)
                    continue
                except (OSError, ValueError):
                    pass
            shutil.rmtree(directory)
            removed.append(directory)
        except OSError:
            # Cleanup is best effort and must not prevent a new session.
            continue
    return tuple(removed)


class SessionIndex:
    """SQLite repository for one application session."""

    def __init__(
        self,
        database_path: str | Path | None = None,
        *,
        temp_root: str | Path | None = None,
        cleanup_on_close: bool = True,
        cleanup_stale: bool = True,
    ) -> None:
        self._owns_directory = database_path is None
        if database_path is None:
            root = Path(temp_root) if temp_root is not None else Path(tempfile.gettempdir())
            root.mkdir(parents=True, exist_ok=True)
            if cleanup_stale:
                cleanup_stale_sessions(root)
            self.session_dir = Path(tempfile.mkdtemp(prefix="excel-search-", dir=root))
            self.database_path = self.session_dir / "index.sqlite3"
        else:
            self.database_path = Path(database_path).resolve()
            self.session_dir = self.database_path.parent
            self.session_dir.mkdir(parents=True, exist_ok=True)
        self.cleanup_on_close = cleanup_on_close
        self._closed = False
        self._lock = threading.RLock()
        self._active_file_id: int | None = None
        self._active_path: Path | None = None
        self._conn = sqlite3.connect(str(self.database_path), check_same_thread=False)
        self._conn.row_factory = sqlite3.Row
        self._conn.executescript(SCHEMA)
        self._conn.execute(
            "INSERT OR REPLACE INTO session_metadata(key, value) VALUES (?, ?)",
            ("created_pid", str(os.getpid())),
        )
        self._conn.commit()
        self._marker: Path | None = self.session_dir / ".active" if self._owns_directory else None
        if self._marker is not None:
            try:
                self._marker.write_text(str(os.getpid()), encoding="ascii")
            except OSError:
                self._marker = None

    @property
    def connection(self) -> sqlite3.Connection:
        return self._conn

    def __enter__(self):
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        self.close()

    def close(self) -> None:
        with self._lock:
            if self._closed:
                return
            if self._conn.in_transaction:
                self._conn.rollback()
            self._conn.close()
            self._closed = True
            marker = self._marker
            if marker is not None:
                try:
                    marker.unlink()
                except OSError:
                    pass
            if self.cleanup_on_close and self._owns_directory:
                try:
                    shutil.rmtree(self.session_dir)
                except OSError:
                    pass

    def begin_file(self, path: str | Path, *, file_order: int = 0) -> int:
        """Start a replace-style transaction for one source workbook."""
        source = Path(path).resolve()
        with self._lock:
            self._ensure_open()
            if self._conn.in_transaction:
                raise IndexTransactionError("another file transaction is active")
            self._conn.execute("DELETE FROM files WHERE path = ?", (str(source),))
            size, mtime_ns = _file_metadata(source)
            cursor = self._conn.execute(
                """INSERT INTO files(path,size,mtime_ns,status,file_order,error,created_at,updated_at)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (
                    str(source),
                    size,
                    mtime_ns,
                    "preparing",
                    file_order,
                    None,
                    time.time(),
                    time.time(),
                ),
            )
            self._active_file_id = int(cursor.lastrowid)
            self._active_path = source
            return self._active_file_id

    def write_chunk(self, chunk: WorksheetChunk) -> None:
        """Insert one bounded worksheet chunk into the active transaction."""
        with self._lock:
            self._ensure_open()
            expected = Path(chunk.file_path).resolve()
            if self._active_file_id is None or self._active_path != expected:
                raise IndexTransactionError("no matching file transaction is active")
            worksheet = chunk.worksheet
            self._conn.execute(
                """INSERT INTO worksheets(file_id,name,sheet_index,visible,first_row,last_row,
                   first_column,last_column) VALUES (?,?,?,?,?,?,?,?)
                   ON CONFLICT(file_id,sheet_index) DO UPDATE SET
                   name=excluded.name, visible=excluded.visible, first_row=excluded.first_row,
                   last_row=excluded.last_row, first_column=excluded.first_column,
                   last_column=excluded.last_column""",
                (
                    self._active_file_id,
                    worksheet.name,
                    worksheet.index,
                    int(worksheet.visible),
                    worksheet.first_row,
                    worksheet.last_row,
                    worksheet.first_column,
                    worksheet.last_column,
                ),
            )
            worksheet_id = int(
                self._conn.execute(
                    "SELECT id FROM worksheets WHERE file_id = ? AND sheet_index = ?",
                    (self._active_file_id, worksheet.index),
                ).fetchone()[0]
            )
            row_numbers = sorted({cell.row for cell in chunk.cells})
            self._conn.executemany(
                "INSERT OR IGNORE INTO physical_rows(worksheet_id,row_number) VALUES (?,?)",
                ((worksheet_id, row_number) for row_number in row_numbers),
            )
            row_ids = {
                int(row["row_number"]): int(row["id"])
                for row in self._conn.execute(
                    """SELECT id,row_number FROM physical_rows
                       WHERE worksheet_id=? AND row_number BETWEEN ? AND ?""",
                    (worksheet_id, chunk.start_row, chunk.end_row),
                ).fetchall()
            }
            cell_values = []
            for cell in chunk.cells:
                row_id = row_ids.get(cell.row)
                if row_id is None:
                    raise IndexTransactionError("physical row could not be created")
                payload = _serialize_value(cell.value)
                cell_values.append(
                    (
                        self._active_file_id,
                        worksheet_id,
                        row_id,
                        cell.row,
                        cell.column,
                        _display_value(cell.value),
                        payload[0],
                        payload[1],
                        _display_value(cell.value),
                        int(cell.is_formula),
                        cell.formula,
                    )
                )
            self._conn.executemany(
                """INSERT OR REPLACE INTO cells(file_id,worksheet_id,row_id,row_number,column_number,
                   value,value_json,value_type,normalized_value,is_formula,formula)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                cell_values,
            )

    # Short aliases make the repository convenient for callback-based callers.
    add_chunk = write_chunk
    insert_chunk = write_chunk

    def complete_file(self, path: str | Path | None = None) -> IndexedFile:
        with self._lock:
            file_id = self._require_active(path)
            self._conn.execute(
                "UPDATE files SET status='ready', error=NULL, updated_at=? WHERE id=?",
                (time.time(), file_id),
            )
            self._conn.commit()
            self._clear_active()
            return self.get_file(file_id)  # type: ignore[return-value]

    def fail_file(
        self, path: str | Path | None, error: BaseException, *, cancelled: bool = False
    ) -> IndexedFile:
        """Rollback all chunks and retain a non-searchable diagnostic row."""
        with self._lock:
            file_id = self._require_active(path)
            source = self._active_path
            order_row = self._conn.execute(
                "SELECT file_order FROM files WHERE id=?", (file_id,)
            ).fetchone()
            order = int(order_row[0]) if order_row else 0
            self._conn.rollback()
            self._clear_active()
            # Preserve a status row while ensuring no child records survive.
            self._conn.execute("DELETE FROM files WHERE path=?", (str(source),))
            size, mtime_ns = _file_metadata(source) if source is not None else (None, None)
            cursor = self._conn.execute(
                """INSERT INTO files(path,size,mtime_ns,status,file_order,error,created_at,updated_at)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (
                    str(source),
                    size,
                    mtime_ns,
                    "cancelled" if cancelled else "failed",
                    order,
                    _error_text(error),
                    time.time(),
                    time.time(),
                ),
            )
            self._conn.commit()
            return self.get_file(int(cursor.lastrowid))  # type: ignore[return-value]

    def remove_file(self, path_or_id: str | Path | int) -> bool:
        with self._lock:
            self._ensure_open()
            if self._conn.in_transaction:
                self._conn.rollback()
                self._clear_active()
            if isinstance(path_or_id, int):
                cursor = self._conn.execute("DELETE FROM files WHERE id=?", (path_or_id,))
            else:
                cursor = self._conn.execute(
                    "DELETE FROM files WHERE path=?", (str(Path(path_or_id).resolve()),)
                )
            self._conn.commit()
            return cursor.rowcount > 0

    delete_file = remove_file

    def clear(self) -> None:
        with self._lock:
            self._ensure_open()
            self._active_file_id = None
            self._active_path = None
            self._conn.rollback()
            self._conn.execute("DELETE FROM files")
            self._conn.commit()

    clear_session = clear

    def get_file(self, file_id: int) -> IndexedFile | None:
        row = self._conn.execute("SELECT * FROM files WHERE id=?", (file_id,)).fetchone()
        return _file_from_row(row) if row else None

    def get_file_by_path(self, path: str | Path) -> IndexedFile | None:
        row = self._conn.execute(
            "SELECT * FROM files WHERE path=?", (str(Path(path).resolve()),)
        ).fetchone()
        return _file_from_row(row) if row else None

    def list_files(self) -> tuple[IndexedFile, ...]:
        rows = self._conn.execute("SELECT * FROM files ORDER BY file_order, id").fetchall()
        return tuple(_file_from_row(row) for row in rows)

    files = list_files

    def searchable_files(self) -> tuple[IndexedFile, ...]:
        rows = self._conn.execute(
            "SELECT * FROM files WHERE status='ready' ORDER BY file_order,id"
        ).fetchall()
        return tuple(_file_from_row(row) for row in rows)

    def list_worksheets(self, file_id: int | None = None) -> tuple[IndexedWorksheet, ...]:
        if file_id is None:
            rows = self._conn.execute(
                "SELECT * FROM worksheets ORDER BY file_id,sheet_index"
            ).fetchall()
        else:
            rows = self._conn.execute(
                "SELECT * FROM worksheets WHERE file_id=? ORDER BY sheet_index", (file_id,)
            ).fetchall()
        return tuple(_worksheet_from_row(row) for row in rows)

    def list_cells(
        self,
        *,
        file_id: int | None = None,
        worksheet_id: int | None = None,
        include_non_searchable: bool = False,
    ) -> tuple[IndexedCell, ...]:
        sql = "SELECT cells.* FROM cells JOIN files ON files.id=cells.file_id WHERE 1=1"
        args: list[Any] = []
        if not include_non_searchable:
            sql += " AND files.status='ready'"
        if file_id is not None:
            sql += " AND file_id=?"
            args.append(file_id)
        if worksheet_id is not None:
            sql += " AND worksheet_id=?"
            args.append(worksheet_id)
        sql += " ORDER BY file_id,worksheet_id,row_number,column_number"
        return tuple(_cell_from_row(row) for row in self._conn.execute(sql, args).fetchall())

    def row_cells(
        self, worksheet_id: int, row_number: int, *, include_non_searchable: bool = False
    ) -> tuple[IndexedCell, ...]:
        status_clause = "" if include_non_searchable else " AND files.status='ready'"
        rows = self._conn.execute(
            """SELECT cells.* FROM cells JOIN files ON files.id=cells.file_id
               WHERE worksheet_id=? AND row_number=?"""
            + status_clause
            + " ORDER BY column_number",
            (worksheet_id, row_number),
        ).fetchall()
        return tuple(_cell_from_row(row) for row in rows)

    def mark_stale(self) -> tuple[IndexedFile, ...]:
        """Mark ready files whose size or nanosecond mtime changed (or vanished)."""
        stale: list[IndexedFile] = []
        for item in self.searchable_files():
            size, mtime_ns = _file_metadata(item.path)
            if size is None or mtime_ns is None or size != item.size or mtime_ns != item.mtime_ns:
                self._conn.execute(
                    "UPDATE files SET status='stale', updated_at=? WHERE id=?",
                    (time.time(), item.id),
                )
        self._conn.commit()
        for item in self.list_files():
            if item.status == "stale":
                stale.append(item)
        return tuple(stale)

    refresh_stale = mark_stale

    def is_file_current(self, path_or_id: str | Path | int) -> bool:
        item = (
            self.get_file(path_or_id)
            if isinstance(path_or_id, int)
            else self.get_file_by_path(path_or_id)
        )
        if item is None or item.status != FileStatus.READY:
            return False
        size, mtime_ns = _file_metadata(item.path)
        return (
            size is not None
            and mtime_ns is not None
            and (size, mtime_ns) == (item.size, item.mtime_ns)
        )

    def _require_active(self, path: str | Path | None) -> int:
        self._ensure_open()
        if self._active_file_id is None:
            raise IndexTransactionError("no file transaction is active")
        if path is not None and Path(path).resolve() != self._active_path:
            raise IndexTransactionError("file transaction path does not match")
        return self._active_file_id

    def _clear_active(self) -> None:
        self._active_file_id = None
        self._active_path = None

    def _ensure_open(self) -> None:
        if self._closed:
            raise IndexErrorBase("index is closed")


class FilePreparationService:
    """Stream :class:`ExcelReader` chunks into a :class:`SessionIndex`."""

    def __init__(
        self,
        index: SessionIndex,
        *,
        reader: ExcelReader | None = None,
        reader_factory: Callable[[], ExcelReader] = ExcelReader,
    ) -> None:
        self.index = index
        self.reader = reader
        self.reader_factory = reader_factory

    def prepare(
        self,
        paths: Iterable[str | Path],
        *,
        cancel_event: threading.Event | None = None,
        progress_callback: Callable[[IngestionProgress], None] | None = None,
    ) -> PreparationResult:
        # Selecting the same path twice should not create two competing source
        # records; preserve the first occurrence for deterministic file order.
        source_paths = tuple(dict.fromkeys(Path(path).resolve() for path in paths))
        reader = self.reader or self.reader_factory()
        outcomes: list[IndexedFile] = []
        file_count = len(source_paths)
        for file_index, path in enumerate(source_paths, start=1):
            if cancel_event is not None and cancel_event.is_set():
                return PreparationResult(tuple(outcomes), cancelled=True)
            self.index.begin_file(path, file_order=file_index - 1)

            def on_progress(
                progress: IngestionProgress, current_file_index: int = file_index
            ) -> None:
                if progress_callback is not None:
                    progress_callback(
                        replace(
                            progress,
                            file_index=current_file_index,
                            file_count=file_count,
                        )
                    )

            try:
                ingestion = reader.read_files(
                    (path,),
                    cancel_event=cancel_event,
                    progress_callback=on_progress,
                    chunk_callback=self.index.write_chunk,
                )
            except IngestionCancelledError as exc:
                outcomes.append(self.index.fail_file(path, exc, cancelled=True))
                return PreparationResult(tuple(outcomes), cancelled=True)
            except Exception as exc:
                outcomes.append(self.index.fail_file(path, exc))
                raise

            result = ingestion.files[0]
            if result.error is None:
                outcomes.append(self.index.complete_file(path))
            else:
                outcomes.append(self.index.fail_file(path, result.error))
        return PreparationResult(tuple(outcomes), cancelled=False)

    prepare_files = prepare

    def prepare_async(
        self,
        paths: Iterable[str | Path],
        *,
        cancel_event: threading.Event | None = None,
        progress_callback: Callable[[IngestionProgress], None] | None = None,
    ) -> Future[PreparationResult]:
        """Prepare files on a dedicated worker thread for UI callers."""
        future: Future[PreparationResult] = Future()
        source_paths = tuple(paths)
        event = cancel_event or threading.Event()

        def worker() -> None:
            if future.set_running_or_notify_cancel():
                try:
                    future.set_result(
                        self.prepare(
                            source_paths,
                            cancel_event=event,
                            progress_callback=progress_callback,
                        )
                    )
                except Exception as exc:  # noqa: BLE001 - worker errors belong on the future
                    future.set_exception(exc)

        threading.Thread(target=worker, name="file-preparation", daemon=True).start()
        return future


IndexRepository = SessionIndex
SessionIndexStore = SessionIndex
IndexStore = SessionIndex
SQLiteIndex = SessionIndex
FilePreparation = FilePreparationService


def _file_metadata(path: Path | None) -> tuple[int | None, int | None]:
    if path is None:
        return None, None
    try:
        stat = path.stat()
    except OSError:
        return None, None
    return int(stat.st_size), int(stat.st_mtime_ns)


def _display_value(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime, date, datetime_time)):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, bytes):
        return base64.b64encode(value).decode("ascii")
    return str(value)


def _serialize_value(value: Any) -> tuple[str | None, str]:
    value_type = type(value).__name__
    try:
        encoded = json.dumps(
            value,
            ensure_ascii=False,
            default=lambda item: {"__type__": type(item).__name__, "value": str(item)},
        )
    except (TypeError, ValueError):
        encoded = json.dumps(str(value), ensure_ascii=False)
        value_type = "str"
    return encoded, value_type


def _error_text(error: BaseException) -> str:
    message = str(error).strip() or type(error).__name__
    return f"{type(error).__name__}: {message}"[:2000]


def _file_from_row(row: sqlite3.Row) -> IndexedFile:
    return IndexedFile(
        int(row["id"]),
        Path(row["path"]),
        row["size"],
        row["mtime_ns"],
        row["status"],
        int(row["file_order"]),
        row["error"],
    )


def _worksheet_from_row(row: sqlite3.Row) -> IndexedWorksheet:
    return IndexedWorksheet(
        int(row["id"]),
        int(row["file_id"]),
        row["name"],
        int(row["sheet_index"]),
        bool(row["visible"]),
        row["first_row"],
        row["last_row"],
        row["first_column"],
        row["last_column"],
    )


def _cell_from_row(row: sqlite3.Row) -> IndexedCell:
    try:
        value = json.loads(row["value_json"]) if row["value_json"] is not None else None
    except (TypeError, ValueError):
        value = row["value"]
    return IndexedCell(
        int(row["id"]),
        int(row["file_id"]),
        int(row["worksheet_id"]),
        int(row["row_number"]),
        int(row["column_number"]),
        value,
        row["value"],
        bool(row["is_formula"]),
        row["formula"],
    )
