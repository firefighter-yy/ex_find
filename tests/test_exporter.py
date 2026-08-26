import os
import time
from pathlib import Path

import pytest

from ex_transform.exporter import (
    ExportCancelledError,
    ExportService,
    ExportSnapshot,
    cleanup_stale_exports,
)
from ex_transform.search_engine import SearchPage


class FakeExporter:
    def __init__(self, *, fail=False):
        self.fail = fail
        self.seen_target = None

    def export(self, snapshot, target, **kwargs):
        self.seen_target = target
        if self.fail:
            raise RuntimeError("boom")
        Path(target).write_bytes(b"valid export")
        return ("warning",)


def test_snapshot_is_immutable_and_export_is_transactional(tmp_path):
    fake = FakeExporter()
    service = ExportService(index=object(), exporter=fake)
    page = SearchPage((), 0, 0, None)
    snapshot = service.snapshot(page)
    assert isinstance(snapshot, ExportSnapshot)
    assert snapshot.results == ()
    result = service.export(snapshot, tmp_path / "result")
    assert result.path == (tmp_path / "result.xlsx").resolve()
    assert result.result_count == 0
    assert result.warnings == ("warning",)
    assert result.path.read_bytes() == b"valid export"
    assert fake.seen_target != result.path
    assert not fake.seen_target.exists()


def test_failed_export_does_not_publish_final_file(tmp_path):
    target = tmp_path / "result.xlsx"
    fake = FakeExporter(fail=True)
    service = ExportService(index=object(), exporter=fake)
    with pytest.raises(RuntimeError):
        service.export(SearchPage((), 0, 0, None), target)
    assert not target.exists()
    assert fake.seen_target is not None
    assert not fake.seen_target.exists()


def test_cancelled_export_does_not_publish_final_file(tmp_path):
    import threading

    event = threading.Event()
    event.set()
    fake = FakeExporter()
    service = ExportService(index=object(), exporter=fake)
    with pytest.raises(ExportCancelledError):
        service.export(SearchPage((), 0, 0, None), tmp_path / "cancel.xlsx", cancel_event=event)
    assert not (tmp_path / "cancel.xlsx").exists()
    assert fake.seen_target is None


def test_cleanup_removes_only_stale_owned_export_files(tmp_path):
    stale = tmp_path / ".excel-search-export-old.xlsx"
    current = tmp_path / ".excel-search-export-current.xlsx"
    unrelated = tmp_path / ".other-temp.xlsx"
    for path in (stale, current, unrelated):
        path.write_bytes(b"temporary")
    old = time.time() - 100
    os.utime(stale, (old, old))

    assert cleanup_stale_exports(tmp_path, older_than_seconds=10) == (stale,)
    assert current.exists()
    assert unrelated.exists()
