import sys
import types

import pytest

from ex_transform.excel_runtime import ExcelApplication, ExcelUnavailableError


def test_runtime_reports_missing_pywin32(monkeypatch) -> None:
    monkeypatch.setitem(sys.modules, "pythoncom", None)
    monkeypatch.setitem(sys.modules, "win32com", None)
    with pytest.raises(ExcelUnavailableError):
        with ExcelApplication():
            pass


def test_runtime_uses_dispatch_ex_and_quits(monkeypatch) -> None:
    calls = []

    class FakeExcel:
        Version = "16.0"

        def __init__(self):
            self.Workbooks = types.SimpleNamespace(Open=lambda *a, **kw: (a, kw))

        def Quit(self):
            calls.append("quit")

    fake_excel = FakeExcel()
    fake_pythoncom = types.SimpleNamespace(CoInitialize=lambda: calls.append("init"), CoUninitialize=lambda: calls.append("uninit"))
    fake_client = types.SimpleNamespace(DispatchEx=lambda name: (calls.append(name) or fake_excel))
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=fake_client))
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    with ExcelApplication() as runtime:
        runtime.open_readonly("book.xlsx")
    assert calls == ["init", "Excel.Application", "quit", "uninit"]
