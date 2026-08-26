"""Repeatable Excel-COM prototypes for the export fidelity spike."""
from __future__ import annotations

import os
import tempfile
from collections import defaultdict
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any, Iterable

from .excel_runtime import ExcelApplication


class ExportMode(str, Enum):
    EXTRACT = "extract"
    APPEARANCE = "appearance"


@dataclass(frozen=True, slots=True)
class ExportRow:
    file_path: Path
    sheet_name: str
    row: int
    first_column: int = 1
    last_column: int | None = None


@dataclass(frozen=True, slots=True)
class ExportResult:
    path: Path
    mode: ExportMode
    rows: int
    warnings: tuple[str, ...] = ()


class ExportFidelitySpike:
    """Compare compact row extraction with full-sheet masking."""

    def __init__(self, application_factory: Any = ExcelApplication) -> None:
        self.application_factory = application_factory

    def export(self, rows: Iterable[ExportRow], destination: str | Path, *, mode: ExportMode = ExportMode.EXTRACT) -> ExportResult:
        selected = tuple(rows)
        if not selected:
            raise ValueError("at least one selected row is required")
        target = Path(destination).resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        suffix = target.suffix.lower() or ".xlsx"
        handle, temporary_name = tempfile.mkstemp(prefix=f".{target.stem}-", suffix=suffix, dir=target.parent)
        os.close(handle)
        temporary = Path(temporary_name)
        output = None
        try:
            with self.application_factory() as runtime:
                books = self._open_sources(runtime, selected)
                try:
                    output = runtime.application.Workbooks.Add()
                    if mode is ExportMode.EXTRACT:
                        self._extract(books, selected, output)
                        warnings = ()
                    elif mode is ExportMode.APPEARANCE:
                        self._appearance(books, selected, output)
                        warnings = ("原貌模式保留未命中数据；结果文件可能包含敏感信息。",)
                    else:
                        raise ValueError(f"unsupported export mode: {mode}")
                    output.SaveAs(str(temporary), FileFormat={".xlsx": 51, ".xlsm": 52}.get(suffix, 51))
                    output.Close(SaveChanges=False)
                    output = None
                finally:
                    for workbook in books.values():
                        runtime.close_workbook(workbook)
            os.replace(temporary, target)
            return ExportResult(target, mode, len(selected), warnings)
        except Exception:
            if output is not None:
                try:
                    output.Close(SaveChanges=False)
                except Exception:
                    pass
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
            raise

    @staticmethod
    def _open_sources(runtime: Any, rows: tuple[ExportRow, ...]) -> dict[Path, Any]:
        books: dict[Path, Any] = {}
        for row in rows:
            path = row.file_path.resolve()
            if path not in books:
                books[path] = runtime.open_readonly(path)
        return books

    @staticmethod
    def _remove_default_sheets(workbook: Any) -> None:
        while workbook.Worksheets.Count > 1:
            workbook.Worksheets(1).Delete()

    def _extract(self, books: dict[Path, Any], rows: tuple[ExportRow, ...], output: Any) -> None:
        groups: dict[tuple[Path, str], list[ExportRow]] = defaultdict(list)
        for row in rows:
            groups[(row.file_path.resolve(), row.sheet_name)].append(row)
        self._remove_default_sheets(output)
        for (path, sheet_name), selected in groups.items():
            source = books[path].Worksheets(sheet_name)
            destination = output.Worksheets(1) if output.Worksheets.Count == 1 else output.Worksheets.Add(After=output.Worksheets(output.Worksheets.Count))
            destination.Name = _sheet_name(path.stem, sheet_name, output)
            for output_row, selected_row in enumerate(sorted(selected, key=lambda item: item.row), 1):
                last_column = selected_row.last_column or _last_column(source, selected_row.row)
                source.Range(source.Cells(selected_row.row, selected_row.first_column), source.Cells(selected_row.row, last_column)).Copy(Destination=destination.Cells(output_row, 1))

    def _appearance(self, books: dict[Path, Any], rows: tuple[ExportRow, ...], output: Any) -> None:
        groups: dict[tuple[Path, str], set[int]] = defaultdict(set)
        for row in rows:
            groups[(row.file_path.resolve(), row.sheet_name)].add(row.row)
        self._remove_default_sheets(output)
        for (path, sheet_name), matched_rows in groups.items():
            source = books[path].Worksheets(sheet_name)
            source.Copy(After=output.Worksheets(output.Worksheets.Count))
            copied = output.Worksheets(output.Worksheets.Count)
            copied.Name = _sheet_name(path.stem, sheet_name, output)
            used = copied.UsedRange
            first = int(used.Row)
            last = first + int(used.Rows.Count) - 1
            for row_number in range(first, last + 1):
                if row_number not in matched_rows:
                    copied.Rows(row_number).Hidden = True


def _last_column(sheet: Any, row: int) -> int:
    return int(sheet.Cells(row, sheet.Columns.Count).End(-4159).Column)


def _sheet_name(stem: str, name: str, workbook: Any) -> str:
    base = " - ".join(part for part in (stem, name) if part)
    safe = "".join("_" if char in "[]:*?/\\" else char for char in base)[:31] or "Result"
    candidate = safe
    index = 2
    existing = {str(workbook.Worksheets(i).Name) for i in range(1, workbook.Worksheets.Count + 1)}
    while candidate in existing:
        suffix = f" ({index})"
        candidate = f"{safe[:31 - len(suffix)]}{suffix}"
        index += 1
    return candidate




