from ex_transform.diagnostics import EnvironmentReport, collect_environment_report, format_report


def test_non_windows_report_does_not_attempt_excel_probe(monkeypatch) -> None:
    monkeypatch.setattr("ex_transform.diagnostics.platform.system", lambda: "Linux")
    report = collect_environment_report()
    assert not report.ready_for_excel
    assert any("requires Windows" in message for message in report.messages)


def test_report_format_is_human_readable() -> None:
    output = format_report(EnvironmentReport("Windows", "3.12.0", True, True, "16.0"))
    assert "Operating system: Windows" in output
    assert "Excel version: 16.0" in output
