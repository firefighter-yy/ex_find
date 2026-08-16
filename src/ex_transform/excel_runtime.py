"""Small, explicit lifecycle wrapper around a private Excel COM instance."""

from __future__ import annotations

from pathlib import Path
from typing import Any


class ExcelUnavailableError(RuntimeError):
    """Raised when Excel COM cannot be initialized on this machine."""


class ExcelApplication:
    """Owns one Excel.Application created with DispatchEx and nothing else."""

    def __init__(self) -> None:
        self._pythoncom: Any = None
        self._excel: Any = None
        self._initialized = False
        self._workbooks: list[Any] = []

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
            self._excel = win32com.client.DispatchEx("Excel.Application")
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
        if self._initialized and self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
            self._initialized = False

    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> None:
        self.close()
