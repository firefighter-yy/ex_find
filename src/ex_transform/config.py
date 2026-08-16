"""Application configuration with no side effects at import time."""

from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class AppConfig:
    """Runtime settings that are safe to pass between UI and worker layers."""

    app_name: str = "Excel Information Search"
    temp_root: Path | None = None
    log_level: str = "INFO"

    @classmethod
    def from_environment(cls) -> "AppConfig":
        configured_root = os.environ.get("EXCEL_SEARCH_TEMP_ROOT")
        return cls(
            temp_root=Path(configured_root) if configured_root else None,
            log_level=os.environ.get("EXCEL_SEARCH_LOG_LEVEL", "INFO").upper(),
        )

    def create_session_directory(self) -> Path:
        """Create a private per-run directory for an index and return its path."""
        parent = self.temp_root
        if parent is not None:
            parent.mkdir(parents=True, exist_ok=True)
            return Path(tempfile.mkdtemp(prefix="excel-search-", dir=parent))
        return Path(tempfile.mkdtemp(prefix="excel-search-"))
