"""Small, explicit lifecycle wrapper around a private Excel COM instance."""

from __future__ import annotations

import gc
import subprocess
import time
from pathlib import Path
from typing import Any


class ExcelUnavailableError(RuntimeError):
    """Raised when Excel COM cannot be initialized on this machine."""


class ExcelApplication:
    """Owns one private Excel.Application and never quits user-owned instances."""

    def __init__(self) -> None:
        self._pythoncom: Any = None
        self._excel: Any = None
        self._initialized = False
        self._workbooks: list[Any] = []
        self._bootstrap_process: Any = None
        self._pid: int | None = None

    def __enter__(self) -> "ExcelApplication":
        try:
            import pythoncom  # type: ignore[import-not-found]
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise ExcelUnavailableError("pywin32 is required for Excel automation") from exc
        try:
            pythoncom.CoInitialize()
            self._pythoncom = pythoncom
            self._initialized = True
            self._excel = self._create_private_instance(win32com.client)
            self._pid = self._application_pid(self._excel)
            self._excel.Visible = False
            self._excel.DisplayAlerts = False
            # 3 = msoAutomationSecurityForceDisable; keep this defensive for mocks/old Office.
            try:
                self._excel.AutomationSecurity = 3
            except Exception:
                pass
            return self
        except Exception as exc:
            self.close()
            raise ExcelUnavailableError("Microsoft Excel could not be started") from exc

    @property
    def application(self) -> Any:
        if self._excel is None:
            raise ExcelUnavailableError("Excel application is not active")
        return self._excel

    def open_readonly(self, path: str | Path) -> Any:
        """Open a workbook without updating links, macros, or the source file."""
        source = str(Path(path).resolve())
        workbook = self.application.Workbooks.Open(
            source,
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        self._workbooks.append(workbook)
        return workbook

    def close_workbook(self, workbook: Any) -> None:
        """Close one workbook without affecting other workbooks in this instance."""
        try:
            workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            self._workbooks.remove(workbook)
        except ValueError:
            pass

    def close(self) -> None:
        for workbook in reversed(self._workbooks):
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        self._workbooks.clear()
        if self._excel is not None:
            try:
                self._excel.Quit()
            except Exception:
                pass
            self._excel = None
        self._terminate_bootstrap()
        self._pid = None
        if self._initialized and self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
            self._initialized = False

    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> None:
        self.close()

    def _create_private_instance(self, client: Any) -> Any:
        """Create an Excel instance without attaching to an existing workbook.

        Excel's COM server can route ``DispatchEx`` to an existing instance when
        that instance already has a workbook open.  A short-lived ``/x`` launch
        forces Excel to register a separate server before the second dispatch.
        The bootstrap process is terminated after the independent COM object is
        acquired; it is never quit through a user-owned COM reference.
        """
        existing_pids = _excel_window_pids()
        excel = client.DispatchEx("Excel.Application")
        pid = self._application_pid(excel)
        if pid is None or pid not in existing_pids:
            return excel

        try:
            self._bootstrap_process = _launch_excel_bootstrap()
            del excel
            gc.collect()
            time.sleep(0.5)
            for _ in range(10):
                candidate = client.DispatchEx("Excel.Application")
                candidate_pid = self._application_pid(candidate)
                if candidate_pid is None or candidate_pid not in existing_pids:
                    return candidate
                del candidate
                gc.collect()
                time.sleep(0.25)
        except Exception:
            self._terminate_bootstrap()
            raise
        self._terminate_bootstrap()
        raise ExcelUnavailableError("Microsoft Excel could not create a private instance")

    @staticmethod
    def _application_pid(excel: Any) -> int | None:
        try:
            import win32process  # type: ignore[import-not-found]

            return int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
        except Exception:
            return None

    def _terminate_bootstrap(self) -> None:
        process = self._bootstrap_process
        self._bootstrap_process = None
        if process is None:
            return
        try:
            if process.poll() is None:
                process.terminate()
                process.wait(timeout=5)
        except Exception:
            try:
                process.kill()
            except Exception:
                pass


def _excel_window_pids() -> set[int]:
    """Return PIDs owning Excel main windows, without importing Excel COM."""
    try:
        import win32gui  # type: ignore[import-not-found]
        import win32process  # type: ignore[import-not-found]
    except ImportError:
        return set()
    pids: set[int] = set()

    def collect(hwnd: int, _extra: Any) -> None:
        try:
            if win32gui.GetClassName(hwnd).upper().startswith("XLMAIN"):
                pids.add(int(win32process.GetWindowThreadProcessId(hwnd)[1]))
        except Exception:
            pass

    try:
        win32gui.EnumWindows(collect, None)
    except Exception:
        return set()
    return pids


def _launch_excel_bootstrap() -> subprocess.Popen[Any]:
    """Launch a hidden ``/x`` Excel process to force a fresh COM server."""
    executable = _excel_executable()
    startup_info = subprocess.STARTUPINFO()
    startup_info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startup_info.wShowWindow = subprocess.SW_HIDE
    return subprocess.Popen(
        [str(executable), "/x", "/e"],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        startupinfo=startup_info,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def _excel_executable() -> Path:
    candidates = (
        Path(r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"),
        Path(r"C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE"),
    )
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise ExcelUnavailableError("Microsoft Excel executable could not be located")


def __getattr__(name: str) -> Any:
    """Lazily expose the reader from the historical runtime module path."""
    if name in {
        "CellRecord",
        "ExcelComReader",
        "ExcelIngestion",
        "ExcelReader",
        "FileReadResult",
        "IngestionCancelledError",
        "IngestionProgress",
        "IngestionResult",
        "WorksheetChunk",
        "WorksheetInfo",
    }:
        from . import excel_ingestion

        return getattr(excel_ingestion, name)
    raise AttributeError(name)
