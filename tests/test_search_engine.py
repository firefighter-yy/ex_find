from datetime import datetime
import threading

from ex_transform import (
    FormulaMode,
    MatchMode,
    SearchService,
    normalize_display_value,
)
from ex_transform.excel_ingestion import CellRecord, WorksheetChunk, WorksheetInfo
from ex_transform.session_index import SessionIndex


def make_chunk(path, rows):
    info = WorksheetInfo("Data", 1, True, 1, len(rows), 1, 4)
    cells = tuple(CellRecord(path, "Data", 1, row_number, column, value, formula is not None, formula)
                  for row_number, row in enumerate(rows, 1)
                  for column, (value, formula) in enumerate(row, 1) if value is not None)
    return WorksheetChunk(path, info, 1, len(rows), 1, 4, cells)


def indexed(tmp_path):
    path = tmp_path / "book.xlsx"
    path.write_bytes(b"book")
    index = SessionIndex(tmp_path / "index.sqlite3", cleanup_on_close=False)
    index.begin_file(path, file_order=0)
    index.write_chunk(make_chunk(path, [
        [("Alpha", None), ("Beta", None), (True, None), (2, "=A1+1")],
        [("alphabet", None), ("other", None), (False, None), (2, None)],
        [("ＡＢＣ\n123", None), (datetime(2024, 1, 2, 3, 4), None), ("00123", None), ("#N/A", None)],
    ]))
    index.complete_file(path)
    return index


def test_normalization_matrix():
    assert normalize_display_value("  Hello\n\u200b WORLD  ") == "hello world"
    assert normalize_display_value("ＡＢＣ") == "abc"
    assert normalize_display_value(1.0) == "1"
    assert normalize_display_value(True) == "true"
    assert normalize_display_value(datetime(2024, 1, 2, 3, 4)) == "2024-01-02 03:04:00"
    assert normalize_display_value("00123") == "00123"


def test_search_and_or_exact_and_deduplicates(tmp_path):
    index = indexed(tmp_path)
    try:
        service = SearchService(index)
        assert [r.row for r in service.search(["alpha", "beta"]).results] == [1]
        assert [r.row for r in service.search(["alpha", "other"], operator="OR").results] == [1, 2]
        assert [r.row for r in service.search("alpha", match=MatchMode.EXACT).results] == [1]
        assert service.search(["alpha", "alpha", " "]).total_count == 2
        assert service.search("*").total_count == 0
    finally:
        index.close()


def test_formula_modes_and_pagination(tmp_path):
    index = indexed(tmp_path)
    try:
        service = SearchService(index)
        assert service.search("a1+1").total_count == 0
        result = service.search("a1+1", formula_mode=FormulaMode.TEXT)
        assert [r.row for r in result.results] == [1]
        page = service.search("alph", offset=1, limit=1)
        assert page.total_count == 2
        assert page.has_more is False
        assert page.results[0].row == 2
        assert page.results[0].matched_columns == (1,)
    finally:
        index.close()


def test_cancelled_search_is_reported(tmp_path):
    index = indexed(tmp_path)
    try:
        event = threading.Event()
        event.set()
        page = SearchService(index).search("alpha", cancel_event=event)
        assert page.cancelled is True
        assert page.results == ()
    finally:
        index.close()





