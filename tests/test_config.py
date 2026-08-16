import tempfile
from pathlib import Path

from ex_transform.config import AppConfig


def test_session_directory_is_created_under_configured_root() -> None:
    # Avoid pytest's tmp_path fixture: some managed Windows profiles deny
    # directory enumeration below the default pytest temp root.
    with tempfile.TemporaryDirectory(prefix="excel-search-test-") as root:
        tmp_path = Path(root)
        session = AppConfig(temp_root=tmp_path).create_session_directory()
        assert session.parent == tmp_path
        assert session.is_dir()
        assert session.name.startswith("excel-search-")
