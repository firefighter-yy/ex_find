import sqlite3
import threading
from pathlib import Path

from ex_transform.excel_ingestion import (
    CellRecord,
    FileReadResult,
    IngestionCancelledError,
    IngestionProgress,
    IngestionResult,
    WorksheetChunk,
    WorksheetInfo,
)
from ex_transform.session_index import (
    FilePreparationService,
    SessionIndex,
    cleanup_stale_sessions,
)


def make_chunk(path: Path, *, sheet_name: str = "Data", sheet_index: int = 1):
    worksheet = WorksheetInfo(sheet_name, sheet_index, True, 2, 3, 1, 2)
    return WorksheetChunk(
        path,
        worksheet,
        2,
        3,
        1,
        2,
        (
            CellRecord(path, sheet_name, sheet_index, 2, 1, "alpha"),
            CellRecord(path, sheet_name, sheet_index, 2, 2, 42),
            CellRecord(path, sheet_name, sheet_index, 3, 1, True, True, "=1=1"),
        ),
    )


def test_schema_and_same_named_sources_are_isolated(tmp_path):
    first = tmp_path / "a" / "book.xlsx"
    second = tmp_path / "b" / "book.xlsx"
    first.parent.mkdir()
    second.parent.mkdir()
    first.write_bytes(b"first")
    second.write_bytes(b"second")

    with SessionIndex(tmp_path / "index.sqlite3", cleanup_on_close=False) as index:
        for order, source in enumerate((first, second)):
            index.begin_file(source, file_order=order)
            index.write_chunk(make_chunk(source, sheet_name="Same"))
            index.complete_file(source)

        files = index.searchable_files()
        assert [item.path for item in files] == [first.resolve(), second.resolve()]
        assert len({item.id for item in files}) == 2
        assert len(index.list_worksheets()) == 2
        assert len(index.list_cells()) == 6
        index_names = {
            row[1] for row in index.connection.execute("PRAGMA index_list('cells')").fetchall()
        }
        assert {"idx_cells_exact", "idx_cells_row"} <= index_names


def test_failure_rolls_back_every_child_record(tmp_path):
    source = tmp_path / "book.xlsx"
    source.write_bytes(b"source")

    with SessionIndex(tmp_path / "index.sqlite3", cleanup_on_close=False) as index:
        index.begin_file(source)
        index.write_chunk(make_chunk(source))
        failed = index.fail_file(source, RuntimeError("read failed"))

        assert failed.status == "failed"
        assert "read failed" in (failed.error or "")
        assert index.searchable_files() == ()
        assert index.list_worksheets() == ()
        assert index.list_cells() == ()


def test_duplicate_import_replaces_old_cells_and_remove_cascades(tmp_path):
    source = tmp_path / "book.xlsx"
    source.write_bytes(b"source")

    with SessionIndex(tmp_path / "index.sqlite3", cleanup_on_close=False) as index:
        index.begin_file(source)
        index.write_chunk(make_chunk(source))
        first = index.complete_file(source)

        index.begin_file(source)
        replacement = make_chunk(source)
        index.write_chunk(
            WorksheetChunk(
                source,
                replacement.worksheet,
                2,
                2,
                1,
                1,
                replacement.cells[:1],
            )
        )
        second = index.complete_file(source)

        assert second.id != first.id
        assert len(index.list_files()) == 1
        assert len(index.list_cells()) == 1
        assert index.remove_file(second.id) is True
        assert index.list_files() == ()
        assert index.list_cells() == ()


def test_source_change_marks_ready_file_stale(tmp_path):
    source = tmp_path / "book.xlsx"
    source.write_bytes(b"before")

    with SessionIndex(tmp_path / "index.sqlite3", cleanup_on_close=False) as index:
        index.begin_file(source)
        index.write_chunk(make_chunk(source))
        prepared = index.complete_file(source)
        source.write_bytes(b"after with a different size")

        stale = index.mark_stale()
        assert [item.id for item in stale] == [prepared.id]
        assert index.searchable_files() == ()


def test_owned_session_is_deleted_and_abandoned_session_is_cleaned(tmp_path):
    abandoned = tmp_path / "excel-search-abandoned"
    abandoned.mkdir()
    (abandoned / "index.sqlite3").write_bytes(b"orphan")
    assert cleanup_stale_sessions(tmp_path) == (abandoned,)

    index = SessionIndex(temp_root=tmp_path)
    session_dir = index.session_dir
    assert index.database_path.exists()
    index.close()
    assert not session_dir.exists()


class StreamingReader:
    def __init__(self, *, fail: bool = False, cancel: bool = False):
        self.fail = fail
        self.cancel = cancel

    def read_files(self, paths, *, cancel_event, progress_callback, chunk_callback):
        path = Path(next(iter(paths))).resolve()
        progress_callback(IngestionProgress(path, 1, 1, None, None, 0, 0, 0, "file_start"))
        chunk_callback(make_chunk(path))
        if self.cancel:
            raise IngestionCancelledError("cancelled")
        error = sqlite3.OperationalError("disk full") if self.fail else None
        return IngestionResult((FileReadResult(path, error=error),))


def test_preparation_service_streams_and_isolates_failures(tmp_path):
    good = tmp_path / "good.xlsx"
    bad = tmp_path / "bad.xlsx"
    good.write_bytes(b"good")
    bad.write_bytes(b"bad")

    with SessionIndex(tmp_path / "index.sqlite3", cleanup_on_close=False) as index:
        ready = FilePreparationService(index, reader=StreamingReader()).prepare((good,))
        assert ready.succeeded[0].path == good.resolve()
        assert len(index.list_cells()) == 3

        failed = FilePreparationService(index, reader=StreamingReader(fail=True)).prepare((bad,))
        assert failed.failed[0].path == bad.resolve()
        assert len(index.list_cells()) == 3


def test_preparation_cancellation_rolls_back_current_file(tmp_path):
    source = tmp_path / "book.xlsx"
    source.write_bytes(b"source")
    event = threading.Event()

    with SessionIndex(tmp_path / "index.sqlite3", cleanup_on_close=False) as index:
        result = FilePreparationService(index, reader=StreamingReader(cancel=True)).prepare(
            (source,), cancel_event=event
        )
        assert result.cancelled is True
        assert result.files[0].status == "cancelled"
        assert index.list_cells() == ()
