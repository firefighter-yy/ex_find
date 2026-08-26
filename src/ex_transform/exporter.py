"""Transactional export of search results to a new Excel workbook."""
from __future__ import annotations
import os
import re
import tempfile
import threading
import time
from collections import defaultdict
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol
from .excel_runtime import ExcelApplication
from .search_engine import SearchPage, SearchResult
from .session_index import SessionIndex

class ExportError(RuntimeError):
    """Raised when a result export cannot be completed safely."""
class ExportCancelledError(ExportError):
    """Raised when the caller cancels an export."""

@dataclass(frozen=True, slots=True)
class ExportSnapshot:
    results: tuple[SearchResult, ...]
    @classmethod
    def from_page(cls, page: SearchPage) -> "ExportSnapshot":
        return cls(tuple(page.results))
    @property
    def result_count(self) -> int:
        return len(self.results)
    @property
    def source_paths(self) -> tuple[Path, ...]:
        return tuple(dict.fromkeys(result.file_path.resolve() for result in self.results))

@dataclass(frozen=True, slots=True)
class ExportProgress:
    completed: int
    total: int
    path: Path | None = None
    stage: str = "copying"

@dataclass(frozen=True, slots=True)
class ExportResult:
    path: Path
    result_count: int
    warnings: tuple[str, ...] = ()

CancelLike = threading.Event | Callable[[], bool] | None
ProgressCallback = Callable[[ExportProgress], None]


def cleanup_stale_exports(
    directory: str | Path, *, older_than_seconds: float = 24 * 60 * 60
) -> tuple[Path, ...]:
    """Remove abandoned hidden export files created by this application."""
    parent = Path(directory)
    if not parent.exists():
        return ()
    cutoff = time.time() - older_than_seconds
    removed: list[Path] = []
    for candidate in parent.glob(".excel-search-export-*"):
        try:
            if candidate.is_file() and candidate.stat().st_mtime <= cutoff:
                candidate.unlink()
                removed.append(candidate)
        except OSError:
            continue
    return tuple(removed)

class WorkbookExporter(Protocol):
    def export(self, snapshot: ExportSnapshot, target: Path, *, cancel_event: CancelLike = None,
               progress_callback: ProgressCallback | None = None) -> tuple[str, ...]: ...

def _cancelled(event: CancelLike) -> bool:
    if event is None:
        return False
    return event.is_set() if hasattr(event, "is_set") else bool(event())
def _check_cancel(event: CancelLike) -> None:
    if _cancelled(event):
        raise ExportCancelledError("export cancelled")
def _metadata(path: Path) -> tuple[int, int]:
    try:
        stat = path.stat()
    except OSError as exc:
        raise ExportError(f"source file is unavailable: {path}") from exc
    return stat.st_size, stat.st_mtime_ns

class ExportService:
    """Validate an index snapshot and atomically publish an exported workbook."""
    def __init__(self, index: SessionIndex, exporter: WorkbookExporter | None = None) -> None:
        self.index = index
        self.exporter = exporter or ExcelComExporter()
    def snapshot(self, page: SearchPage) -> ExportSnapshot:
        return ExportSnapshot.from_page(page)
    def export(self, page_or_snapshot: SearchPage | ExportSnapshot, target: str | Path, *,
               cancel_event: CancelLike = None, progress_callback: ProgressCallback | None = None) -> ExportResult:
        snapshot = page_or_snapshot if isinstance(page_or_snapshot, ExportSnapshot) else self.snapshot(page_or_snapshot)
        destination = Path(target).expanduser().resolve()
        if destination.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            destination = destination.with_suffix(".xlsx")
        destination.parent.mkdir(parents=True, exist_ok=True)
        cleanup_stale_exports(destination.parent)
        self._validate_sources(snapshot)
        _check_cancel(cancel_event)
        fd, temporary_name = tempfile.mkstemp(
            prefix=".excel-search-export-", suffix=destination.suffix, dir=destination.parent
        )
        os.close(fd)
        temporary = Path(temporary_name)
        try:
            warnings = self.exporter.export(snapshot, temporary, cancel_event=cancel_event, progress_callback=progress_callback)
            _check_cancel(cancel_event)
            self._validate_sources(snapshot)
            if not temporary.exists() or temporary.stat().st_size == 0:
                raise ExportError("Excel did not create a valid output file")
            os.replace(temporary, destination)
            if progress_callback is not None:
                progress_callback(ExportProgress(snapshot.result_count, snapshot.result_count, destination, "complete"))
            return ExportResult(destination, snapshot.result_count, tuple(warnings))
        except ExportCancelledError:
            raise
        except Exception as exc:
            if isinstance(exc, ExportError):
                raise
            raise ExportError("workbook export failed") from exc
        finally:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
    def _validate_sources(self, snapshot: ExportSnapshot) -> None:
        for path in snapshot.source_paths:
            indexed = self.index.get_file_by_path(path)
            if indexed is None or indexed.status != "ready":
                raise ExportError("源文件尚未准备完成，请重新准备文件")
            size, mtime_ns = _metadata(path)
            if size != indexed.size or mtime_ns != indexed.mtime_ns:
                self.index.mark_stale()
                raise ExportError("源文件已变化，请重新准备文件")

class ExcelComExporter:
    """Copy source worksheets with Excel itself, preserving native content."""
    def __init__(self, application_factory: Callable[[], Any] = ExcelApplication) -> None:
        self.application_factory = application_factory
    def export(self, snapshot: ExportSnapshot, target: Path, *, cancel_event: CancelLike = None,
               progress_callback: ProgressCallback | None = None) -> tuple[str, ...]:
        if not snapshot.results:
            pass
        grouped: dict[tuple[Path, str], list[SearchResult]] = defaultdict(list)
        for result in snapshot.results:
            grouped[(result.file_path.resolve(), result.worksheet_name)].append(result)
        with self.application_factory() as runtime:
            excel = runtime.application
            destination = excel.Workbooks.Add()
            try:
                first = destination.Worksheets(1)
                first.Name = "导出结果"
                self._write_traceability(first, snapshot.results)
                for index, ((source_path, worksheet_name), rows) in enumerate(grouped.items(), 1):
                    _check_cancel(cancel_event)
                    source = runtime.open_readonly(source_path)
                    try:
                        source_sheet = source.Worksheets(worksheet_name)
                        source_sheet.Copy(After=destination.Worksheets(destination.Worksheets.Count))
                        copied = destination.Worksheets(destination.Worksheets.Count)
                        copied.Name = _unique_sheet_name(destination, worksheet_name, index)
                        self._hide_unmatched_rows(copied, {item.row for item in rows})
                        if progress_callback is not None:
                            progress_callback(ExportProgress(index, len(grouped), source_path, "copying"))
                    finally:
                        runtime.close_workbook(source)
                self._save_as(destination, target)
            finally:
                try:
                    destination.Close(SaveChanges=False)
                except Exception:
                    pass
        return ()
    @staticmethod
    def _hide_unmatched_rows(sheet: Any, keep: set[int]) -> None:
        used = sheet.UsedRange
        first = int(used.Row)
        last = first + int(used.Rows.Count) - 1
        for row in range(first, last + 1):
            if row not in keep:
                sheet.Rows(row).Hidden = True
    @staticmethod
    def _write_traceability(sheet: Any, results: Iterable[SearchResult]) -> None:
        values = [("来源文件", "来源工作表", "原始行号", "命中关键词", "导出工作表")]
        values += [(str(item.file_path), item.worksheet_name, item.row, ", ".join(item.matched_keywords), "") for item in results]
        sheet.Range("A1").Resize(len(values), 5).Value = values
    @staticmethod
    def _save_as(workbook: Any, target: Path) -> None:
        file_format = {".xlsx": 51, ".xlsm": 52, ".xls": 56}.get(target.suffix.lower(), 51)
        workbook.SaveAs(str(target), FileFormat=file_format)

_INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")
def _unique_sheet_name(workbook: Any, requested: str, ordinal: int) -> str:
    base = _INVALID_SHEET_CHARS.sub("_", requested).strip() or "结果"
    base = base[:31]
    existing = {str(workbook.Worksheets(i).Name).casefold() for i in range(1, workbook.Worksheets.Count + 1)}
    candidate = base
    suffix = 2
    while candidate.casefold() in existing:
        tail = f"_{suffix}"
        candidate = f"{base[:31-len(tail)]}{tail}"
        suffix += 1
    return candidate
