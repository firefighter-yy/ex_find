"""Environment diagnostics for Windows, Python, Excel, and COM."""

from __future__ import annotations

import platform
import sys
from dataclasses import dataclass, field


@dataclass(slots=True)
class EnvironmentReport:
    operating_system: str
    python_version: str
    excel_available: bool = False
    com_available: bool = False
    excel_version: str | None = None
    messages: list[str] = field(default_factory=list)

    @property
    def ready_for_excel(self) -> bool:
        return self.operating_system == "Windows" and self.com_available and self.excel_available


def _probe_excel(report: EnvironmentReport) -> None:
    """Probe a private Excel instance and always release COM resources."""
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        report.messages.append("pywin32 is not installed; Excel COM is unavailable")
        return

    initialized = False
    excel = None
    try:
        pythoncom.CoInitialize()
        initialized = True
        report.com_available = True
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        report.excel_available = True
        report.excel_version = str(getattr(excel, "Version", "unknown"))
    except Exception as exc:  # COM errors vary by Office version and machine policy.
        report.messages.append(f"Excel COM probe failed: {type(exc).__name__}")
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def collect_environment_report(*, probe_excel: bool = True) -> EnvironmentReport:
    report = EnvironmentReport(
        operating_system=platform.system(),
        python_version=platform.python_version(),
    )
    if report.operating_system != "Windows":
        report.messages.append("Microsoft Excel COM requires Windows")
    elif probe_excel:
        _probe_excel(report)
    return report


def format_report(report: EnvironmentReport) -> str:
    lines = [
        f"Operating system: {report.operating_system}",
        f"Python: {report.python_version}",
        f"COM available: {'yes' if report.com_available else 'no'}",
        f"Excel available: {'yes' if report.excel_available else 'no'}",
    ]
    if report.excel_version:
        lines.append(f"Excel version: {report.excel_version}")
    lines.extend(f"Notice: {message}" for message in report.messages)
    return "\n".join(lines)
