from ex_transform.logging_utils import redact_message


def test_redacts_workbook_paths_and_explicit_values() -> None:
    message = redact_message(r"Opening C:\data\private.xlsx for search", ["secret"])
    assert "private.xlsx" not in message
    assert "[workbook]" in message
    assert redact_message("secret value", ["secret"]) == "[redacted] value"
