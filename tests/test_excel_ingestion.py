import threading
import types
from pathlib import Path

import pytest

from ex_transform.excel_ingestion import ExcelReader, IngestionCancelledError


class FakeRange:
    def __init__(self, values, formulas=None):
        self.Value2 = values
        self.Formula = formulas if formulas is not None else values


class FakeCells:
    def __init__(self, sheet):
        self.sheet = sheet

    def __call__(self, row, column):
        return (row, column)

    def Find(self, **kwargs):
        if kwargs.get("SearchDirection") == 1:
            return types.SimpleNamespace(Row=2, Column=3)
        if kwargs.get("SearchOrder") == 1:
            return types.SimpleNamespace(Row=3, Column=3)
        return types.SimpleNamespace(Row=3, Column=4)


class FakeSheet:
    Name = "Data"
    Visible = -1

    def __init__(self, values, formulas):
        self.Cells = FakeCells(self)
        self.UsedRange = types.SimpleNamespace(Row=1, Column=1)
        self.values = values
        self.formulas = formulas

    def Range(self, start, end):
        return FakeRange(self.values, self.formulas)

    def Rows(self, row):
        return types.SimpleNamespace(Hidden=False)

    def Columns(self, column):
        return types.SimpleNamespace(Hidden=False)


class FakeWorkbook:
    def __init__(self, sheet):
        self.Worksheets = types.SimpleNamespace(Count=1, Item=lambda number: sheet)
        self.closed = False

    def Close(self, **kwargs):
        self.closed = True


class FakeApplication:
    def __init__(self):
        self.workbooks = []

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def open_readonly(self, path):
        if Path(path).name == "broken.xlsx":
            raise OSError("cannot open")
        workbook = FakeWorkbook(
            FakeSheet(
                (("hello", 2), (None, "")),
                (("hello", 2), (None, "=A1+1")),
            )
        )
        self.workbooks.append(workbook)
        return workbook

    def close_workbook(self, workbook):
        workbook.Close(SaveChanges=False)


def test_reader_batches_values_and_isolates_file_failure(tmp_path):
    reader = ExcelReader(chunk_rows=2, application_factory=FakeApplication)
    result = reader.read_files([tmp_path / "good.xlsx", tmp_path / "broken.xlsx"])

    assert len(result.succeeded) == 1
    assert len(result.failed) == 1
    cells = result.succeeded[0].chunks[0].cells
    assert [(cell.row, cell.column, cell.value) for cell in cells] == [
        (2, 3, "hello"),
        (2, 4, 2),
        (3, 4, ""),
    ]
    assert cells[-1].is_formula is True
    assert cells[-1].formula == "=A1+1"


def test_reader_honors_cancellation(tmp_path):
    event = threading.Event()
    event.set()
    with pytest.raises(IngestionCancelledError):
        ExcelReader(application_factory=FakeApplication).read_files(
            [tmp_path / "good.xlsx"], cancel_event=event
        )


def test_reader_async_returns_result(tmp_path):
    result = ExcelReader(application_factory=FakeApplication).read_async(
        [tmp_path / "good.xlsx"]
    ).result(timeout=2)
    assert result.succeeded[0].file_path.name == "good.xlsx"
