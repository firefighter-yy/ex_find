"""Read-only, block-oriented ingestion from Excel COM.

The classes in this module deliberately contain no UI or index-store logic.  COM
objects are used only inside :class:`ExcelReader.read_files` (or its worker
thread), and the values crossing that boundary are immutable dataclasses.
"""

from __future__ import annotations

import threading
from collections.abc import Callable, Iterable, Iterator
from concurrent.futures import Future
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .excel_runtime import ExcelApplication


class IngestionCancelledError(RuntimeError):
    """Raised when a read is stopped through its cancellation event."""


@dataclass(frozen=True, slots=True)
class CellRecord:
    """One non-empty physical cell returned by Excel."""

    file_path: Path
    sheet_name: str
    sheet_index: int
    row: int
    column: int
    value: Any
    is_formula: bool = False
    formula: str | None = None


@dataclass(frozen=True, slots=True)
class WorksheetInfo:
    """Worksheet identity and the effective value/formula bounds."""

    name: str
    index: int
    visible: bool
    first_row: int | None
    last_row: int | None
    first_column: int | None
    last_column: int | None


@dataclass(frozen=True, slots=True)
class WorksheetChunk:
    """A bounded batch of cells from one worksheet."""

    file_path: Path
    worksheet: WorksheetInfo
    start_row: int
    end_row: int
    start_column: int
    end_column: int
    cells: tuple[CellRecord, ...]


@dataclass(frozen=True, slots=True)
class IngestionProgress:
    """Progress event emitted after each worksheet block."""

    file_path: Path
    file_index: int
    file_count: int
    sheet_name: str | None
    sheet_index: int | None
    rows_done: int
    rows_total: int
    blocks_done: int
    stage: str = "reading"


@dataclass(frozen=True, slots=True)
class FileReadResult:
    """Outcome for one file; an error here does not abort other files."""

    file_path: Path
    chunks: tuple[WorksheetChunk, ...] = ()
    error: Exception | None = None

    @property
    def succeeded(self) -> bool:
        return self.error is None


@dataclass(frozen=True, slots=True)
class IngestionResult:
    """All per-file outcomes from one ingestion request."""

    files: tuple[FileReadResult, ...]

    @property
    def succeeded(self) -> tuple[FileReadResult, ...]:
        return tuple(item for item in self.files if item.succeeded)

    @property
    def failed(self) -> tuple[FileReadResult, ...]:
        return tuple(item for item in self.files if not item.succeeded)


ProgressCallback = Callable[[IngestionProgress], None]
ChunkCallback = Callable[[WorksheetChunk], None]
CancelLike = threading.Event | Callable[[], bool] | None


class ExcelReader:
    """Read one or more workbooks using one private Excel instance.

    ``read_files`` is intended for an already dedicated worker thread.  For UI
    callers, ``read_async`` creates exactly one worker thread and initializes
    COM there, returning a standard :class:`~concurrent.futures.Future`.
    """

    def __init__(
        self,
        *,
        chunk_rows: int = 512,
        include_hidden_sheets: bool = True,
        include_hidden_rows: bool = True,
        include_hidden_columns: bool = True,
        application_factory: Callable[[], ExcelApplication] = ExcelApplication,
    ) -> None:
        if chunk_rows < 1:
            raise ValueError("chunk_rows must be positive")
        self.chunk_rows = chunk_rows
        self.include_hidden_sheets = include_hidden_sheets
        self.include_hidden_rows = include_hidden_rows
        self.include_hidden_columns = include_hidden_columns
        self.application_factory = application_factory

    def read_async(
        self,
        paths: Iterable[str | Path],
        *,
        cancel_event: threading.Event | None = None,
        progress_callback: ProgressCallback | None = None,
        chunk_callback: ChunkCallback | None = None,
    ) -> Future[IngestionResult]:
        """Run :meth:`read_files` on a fresh thread with a COM apartment."""
        future: Future[IngestionResult] = Future()
        source_paths = tuple(paths)
        event = cancel_event or threading.Event()

        def worker() -> None:
            if future.set_running_or_notify_cancel():
                try:
                    future.set_result(
                        self.read_files(
                            source_paths,
                            cancel_event=event,
                            progress_callback=progress_callback,
                            chunk_callback=chunk_callback,
                        )
                    )
                except Exception as exc:  # noqa: BLE001 - isolate worker failures
                    future.set_exception(exc)

        threading.Thread(target=worker, name="excel-ingestion", daemon=True).start()
        return future

    def read_files(
        self,
        paths: Iterable[str | Path],
        *,
        cancel_event: CancelLike = None,
        progress_callback: ProgressCallback | None = None,
        chunk_callback: ChunkCallback | None = None,
    ) -> IngestionResult:
        """Read files in order and isolate failures to their individual file."""
        source_paths = tuple(Path(path).resolve() for path in paths)
        outcomes: list[FileReadResult] = []
        with self.application_factory() as application:
            for file_index, path in enumerate(source_paths, start=1):
                self._check_cancel(cancel_event)
                try:
                    if progress_callback is not None:
                        progress_callback(
                            IngestionProgress(
                                path,
                                file_index,
                                len(source_paths),
                                None,
                                None,
                                0,
                                0,
                                0,
                                stage="file_start",
                            )
                        )
                    chunks_list: list[WorksheetChunk] = []
                    for chunk in self._read_file(
                        application,
                        path,
                        file_index=file_index,
                        file_count=len(source_paths),
                        cancel_event=cancel_event,
                        progress_callback=progress_callback,
                    ):
                        if chunk_callback is None:
                            chunks_list.append(chunk)
                        else:
                            chunk_callback(chunk)
                    chunks = tuple(chunks_list)
                    outcomes.append(FileReadResult(path, chunks=chunks))
                    if progress_callback is not None:
                        progress_callback(
                            IngestionProgress(
                                path,
                                file_index,
                                len(source_paths),
                                None,
                                None,
                                0,
                                0,
                                0,
                                stage="file_complete",
                            )
                        )
                except IngestionCancelledError:
                    raise
                except Exception as exc:  # noqa: BLE001 - isolate one file failure
                    outcomes.append(FileReadResult(path, error=exc))
                    if progress_callback is not None:
                        progress_callback(
                            IngestionProgress(
                                path,
                                file_index,
                                len(source_paths),
                                None,
                                None,
                                0,
                                0,
                                0,
                                stage="failed",
                            )
                        )
        return IngestionResult(tuple(outcomes))

    # Friendly alias for service layers that use ``ingest`` terminology.
    ingest = read_files
    read = read_files
    read_workbooks = read_files

    def _read_file(
        self,
        application: ExcelApplication,
        path: Path,
        *,
        file_index: int,
        file_count: int,
        cancel_event: CancelLike,
        progress_callback: ProgressCallback | None,
    ) -> Iterator[WorksheetChunk]:
        workbook = application.open_readonly(path)
        try:
            worksheets = workbook.Worksheets
            count = int(getattr(worksheets, "Count", 0))
            for sheet_number in range(1, count + 1):
                self._check_cancel(cancel_event)
                sheet = worksheets.Item(sheet_number)
                visible = self._sheet_visible(sheet)
                if not visible and not self.include_hidden_sheets:
                    continue
                info, bounds = self._worksheet_bounds(sheet, sheet_number, visible)
                if progress_callback is not None:
                    progress_callback(
                        IngestionProgress(
                            path,
                            file_index,
                            file_count,
                            info.name,
                            sheet_number,
                            0,
                            0,
                            0,
                            stage="worksheet_start",
                        )
                    )
                if bounds is None:
                    if progress_callback is not None:
                        progress_callback(
                            IngestionProgress(
                                path, file_index, file_count, info.name, sheet_number, 0, 0, 0
                            )
                        )
                    continue
                first_row, last_row, first_col, last_col = bounds
                total_rows = last_row - first_row + 1
                rows_done = 0
                for blocks_done, start_row in enumerate(
                    range(first_row, last_row + 1, self.chunk_rows), start=1
                ):
                    self._check_cancel(cancel_event)
                    end_row = min(last_row, start_row + self.chunk_rows - 1)
                    chunk = self._read_chunk(
                        sheet,
                        path,
                        info,
                        start_row,
                        end_row,
                        first_col,
                        last_col,
                        cancel_event,
                    )
                    rows_done += end_row - start_row + 1
                    yield chunk
                    if progress_callback is not None:
                        progress_callback(
                            IngestionProgress(
                                path,
                                file_index,
                                file_count,
                                info.name,
                                sheet_number,
                                rows_done,
                                total_rows,
                                blocks_done,
                            )
                        )
        finally:
            close_workbook = getattr(application, "close_workbook", None)
            if callable(close_workbook):
                close_workbook(workbook)
            else:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:  # noqa: BLE001,S110 - cleanup must not mask the original error
                    pass

    def _worksheet_bounds(
        self, sheet: Any, sheet_number: int, visible: bool
    ) -> tuple[WorksheetInfo, tuple[int, int, int, int] | None]:
        used = sheet.UsedRange
        used_first_row = int(getattr(used, "Row", 1))
        used_first_col = int(getattr(used, "Column", 1))
        first_row_cell = self._find(sheet, by_columns=False, previous=False)
        first_col_cell = self._find(sheet, by_columns=True, previous=False)
        last_row_cell = self._find(sheet, by_columns=False, previous=True)
        last_col_cell = self._find(sheet, by_columns=True, previous=True)
        if last_row_cell is None or last_col_cell is None:
            info = WorksheetInfo(str(sheet.Name), sheet_number, visible, None, None, None, None)
            return info, None
        first_row = (
            int(getattr(first_row_cell, "Row", used_first_row))
            if first_row_cell is not None
            else used_first_row
        )
        first_col = (
            int(getattr(first_col_cell, "Column", used_first_col))
            if first_col_cell is not None
            else used_first_col
        )
        last_row = int(last_row_cell.Row)
        last_col = int(last_col_cell.Column)
        if last_row < first_row or last_col < first_col:
            info = WorksheetInfo(str(sheet.Name), sheet_number, visible, None, None, None, None)
            return info, None
        info = WorksheetInfo(
            str(sheet.Name), sheet_number, visible, first_row, last_row, first_col, last_col
        )
        return info, (first_row, last_row, first_col, last_col)

    def _read_chunk(
        self,
        sheet: Any,
        path: Path,
        info: WorksheetInfo,
        start_row: int,
        end_row: int,
        first_col: int,
        last_col: int,
        cancel_event: CancelLike,
    ) -> WorksheetChunk:
        self._check_cancel(cancel_event)
        cell_range = sheet.Range(sheet.Cells(start_row, first_col), sheet.Cells(end_row, last_col))
        values = _matrix(cell_range.Value2, end_row - start_row + 1, last_col - first_col + 1)
        formulas = _matrix(
            getattr(cell_range, "Formula", None),
            end_row - start_row + 1,
            last_col - first_col + 1,
        )
        formula_flags = _matrix(
            getattr(cell_range, "HasFormula", None),
            end_row - start_row + 1,
            last_col - first_col + 1,
        )
        cells: list[CellRecord] = []
        for row_offset, value_row in enumerate(values):
            row = start_row + row_offset
            if not self.include_hidden_rows and self._hidden(sheet.Rows(row)):
                continue
            for col_offset, value in enumerate(value_row):
                column = first_col + col_offset
                if not self.include_hidden_columns and self._hidden(sheet.Columns(column)):
                    continue
                formula_value = formulas[row_offset][col_offset]
                formula_text = (
                    formula_value
                    if isinstance(formula_value, str) and formula_value.startswith("=")
                    else None
                )
                has_formula = formula_flags[row_offset][col_offset]
                is_formula = (
                    bool(has_formula)
                    if isinstance(has_formula, bool)
                    else formula_text is not None
                )
                if value is None and not is_formula and formula_value in (None, ""):
                    continue
                cells.append(
                    CellRecord(
                        path,
                        info.name,
                        info.index,
                        row,
                        column,
                        value,
                        is_formula,
                        formula_text,
                    )
                )
        return WorksheetChunk(path, info, start_row, end_row, first_col, last_col, tuple(cells))

    @staticmethod
    def _find(sheet: Any, *, by_columns: bool, previous: bool) -> Any | None:
        # Excel constants: xlFormulas=-4123, xlByRows=1, xlByColumns=2,
        # xlPrevious=2.  Numeric constants avoid importing Excel on non-Windows tests.
        try:
            arguments = {
                "What": "*",
                "LookIn": -4123,
                "LookAt": 2,
                "SearchOrder": 2 if by_columns else 1,
                "SearchDirection": 2 if previous else 1,
                "MatchCase": False,
            }
            try:
                arguments["After"] = sheet.Cells(1, 1)
            except Exception:  # noqa: BLE001,S110 - mocked/old Excel may not expose Cells(index)
                pass
            return sheet.Cells.Find(**arguments)
        except Exception:  # noqa: BLE001 - COM failures are represented as an empty effective range
            return None

    @staticmethod
    def _sheet_visible(sheet: Any) -> bool:
        # xlSheetVisible is -1; unknown/mocked objects are treated as visible.
        try:
            value = sheet.Visible
            if isinstance(value, bool):
                return value
            return int(value) == -1
        except Exception:  # noqa: BLE001 - worksheet visibility varies across COM proxies
            return True

    @staticmethod
    def _hidden(item: Any) -> bool:
        try:
            return bool(item.Hidden)
        except Exception:  # noqa: BLE001 - hidden metadata is optional across worksheet proxies
            return False

    @staticmethod
    def _check_cancel(cancel_event: CancelLike) -> None:
        cancelled = (
            cancel_event.is_set()
            if isinstance(cancel_event, threading.Event)
            else bool(cancel_event and cancel_event())
        )
        if cancelled:
            raise IngestionCancelledError("Excel ingestion was cancelled")


def _matrix(value: Any, rows: int, columns: int) -> list[list[Any]]:
    """Convert Excel's scalar/tuple return values to a predictable matrix."""
    if rows == 1 and columns == 1:
        return [[value]]
    if isinstance(value, (tuple, list)):
        if rows == 1 and (not value or not isinstance(value[0], (tuple, list))):
            result = [list(value)]
        else:
            result = [list(row) if isinstance(row, (tuple, list)) else [row] for row in value]
    else:
        result = [[value]]
    result.extend([[] for _ in range(rows - len(result))])
    return [row + [None] * (columns - len(row)) for row in result[:rows]]


# Compatibility names for callers that use the more explicit terminology.
ExcelIngestion = ExcelReader
ExcelComReader = ExcelReader
CellData = CellRecord
