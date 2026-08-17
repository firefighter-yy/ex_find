"""Opt-in integration coverage against a real Microsoft Excel installation."""

from __future__ import annotations

import gc
import hashlib
import os
import threading
from pathlib import Path

import pytest

from ex_transform.excel_ingestion import ExcelReader, IngestionCancelledError
from ex_transform.excel_runtime import ExcelApplication, ExcelUnavailableError

pytestmark = pytest.mark.integration


def _require_excel_integration() -> None:
    if os.environ.get("EXCEL_COM_INTEGRATION") != "1":
        pytest.skip("set EXCEL_COM_INTEGRATION=1 to run real Excel COM tests")


def _digest(path: Path) -> tuple[str, int, int]:
    return (
        hashlib.sha256(path.read_bytes()).hexdigest(),
        path.stat().st_size,
        path.stat().st_mtime_ns,
    )


@pytest.fixture
def excel_samples(tmp_path: Path) -> tuple[Path, Path, Path]:
    _require_excel_integration()
    xlsx = tmp_path / "sample.xlsx"
    xlsm = tmp_path / "sample.xlsm"
    broken = tmp_path / "broken.xlsx"
    try:
        with ExcelApplication() as runtime:
            workbook = runtime.application.Workbooks.Add()
            while workbook.Worksheets.Count < 3:
                workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
            data = workbook.Worksheets(1)
            data.Name = "Data"
            data.Range("A1:C3").Value2 = (
                ("id", "left", "right"),
                (1, 2, 3),
                (2, 5, 7),
            )
            data.Range("D2").Formula = "=SUM(B2:C2)"
            data.Range("D3").Formula = "=SUM(B3:C3)"
            data.Rows(3).Hidden = True
            data.Columns(3).Hidden = True
            data.Range("Z100000").Interior.Color = 65535
            workbook.Worksheets(2).Name = "Empty"
            hidden = workbook.Worksheets(3)
            hidden.Name = "Hidden"
            hidden.Range("A1").Value2 = "secret"
            hidden.Visible = 0
            workbook.SaveAs(str(xlsx), FileFormat=51)
            workbook.SaveAs(str(xlsm), FileFormat=52)
            del data, hidden
            gc.collect()
            runtime.close_workbook(workbook)
            del workbook
    except ExcelUnavailableError as exc:
        pytest.skip(str(exc))
    broken.write_bytes(b"not an Excel workbook")
    return xlsx, xlsm, broken


def test_real_excel_reads_xlsx_xlsm_and_isolates_failure(
    excel_samples: tuple[Path, Path, Path],
) -> None:
    _xlsx, _xlsm, _broken = excel_samples
    before = {path.name: _digest(path) for path in excel_samples}
    result = ExcelReader(chunk_rows=2).read_files(excel_samples)
    after = {path.name: _digest(path) for path in excel_samples}

    assert [item.succeeded for item in result.files] == [True, True, False]
    assert before == after
    cells = [cell for item in result.succeeded for chunk in item.chunks for cell in chunk.cells]
    formulas = [cell for cell in cells if cell.is_formula]
    assert {(cell.value, cell.formula) for cell in formulas} == {
        (5.0, "=SUM(B2:C2)"),
        (12.0, "=SUM(B3:C3)"),
    }
    assert {(cell.file_path.suffix, cell.sheet_name) for cell in cells} >= {
        (".xlsx", "Hidden"),
        (".xlsm", "Hidden"),
    }
    data_chunks = [
        chunk
        for item in result.succeeded
        for chunk in item.chunks
        if chunk.worksheet.name == "Data"
    ]
    assert max(chunk.end_row for chunk in data_chunks) == 3


def test_real_excel_hidden_switches_and_cancellation(
    excel_samples: tuple[Path, Path, Path],
) -> None:
    xlsx, _, _ = excel_samples
    result = ExcelReader(
        chunk_rows=2,
        include_hidden_sheets=False,
        include_hidden_rows=False,
        include_hidden_columns=False,
    ).read_files([xlsx])
    cells = [cell for item in result.succeeded for chunk in item.chunks for cell in chunk.cells]
    assert [(cell.row, cell.column) for cell in cells] == [
        (1, 1),
        (1, 2),
        (2, 1),
        (2, 2),
        (2, 4),
    ]

    cancel = threading.Event()

    def cancel_after_first_chunk(progress) -> None:
        if progress.stage == "reading":
            cancel.set()

    with pytest.raises(IngestionCancelledError):
        ExcelReader(chunk_rows=1).read_files(
            [xlsx], cancel_event=cancel, progress_callback=cancel_after_first_chunk
        )


@pytest.mark.performance
def test_real_excel_streams_one_hundred_thousand_rows(tmp_path: Path) -> None:
    _require_excel_integration()
    if os.environ.get("EXCEL_COM_PERFORMANCE") != "1":
        pytest.skip("set EXCEL_COM_PERFORMANCE=1 to run the 100,000-row benchmark")
    source = tmp_path / "large-100k.xlsx"
    with ExcelApplication() as runtime:
        workbook = runtime.application.Workbooks.Add()
        sheet = workbook.Worksheets(1)
        rows = (("id", "name", "amount"),) + tuple(
            (index, f"row-{index}", index * 1.5) for index in range(1, 100000)
        )
        sheet.Range("A1:C100000").Value2 = rows
        workbook.SaveAs(str(source), FileFormat=51)
        runtime.close_workbook(workbook)

    counts = {"chunks": 0, "cells": 0}

    def collect(chunk) -> None:
        counts["chunks"] += 1
        counts["cells"] += len(chunk.cells)

    before = _digest(source)
    result = ExcelReader(chunk_rows=4096).read_async(
        [source], chunk_callback=collect
    ).result(timeout=120)
    assert len(result.succeeded) == 1
    assert result.succeeded[0].chunks == ()
    assert counts == {"chunks": 25, "cells": 300000}
    assert _digest(source) == before
