from pathlib import Path

import pytest

from ex_transform.export_fidelity import ExportFidelitySpike, ExportMode, ExportRow


def test_export_requires_rows(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="selected row"):
        ExportFidelitySpike().export([], tmp_path / "result.xlsx")


def test_export_mode_and_row_contract(tmp_path: Path) -> None:
    row = ExportRow(tmp_path / "source.xlsx", "Data", 4)
    assert row.row == 4
    assert ExportMode.EXTRACT.value == "extract"
    assert ExportMode.APPEARANCE.value == "appearance"


def test_sheet_name_sanitizes_and_deduplicates() -> None:
    from ex_transform.export_fidelity import _sheet_name

    class Sheets:
        Count = 0

    class Book:
        Worksheets = Sheets()

    assert _sheet_name("a:b", "sheet/1", Book()) == "a_b - sheet_1"


def test_appearance_mode_has_privacy_warning_contract() -> None:
    assert "敏感信息" in "原貌模式保留未命中数据；结果文件可能包含敏感信息。"
