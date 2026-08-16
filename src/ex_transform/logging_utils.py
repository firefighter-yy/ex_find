"""Logging helpers that prevent workbook contents and user input leaking to logs."""

from __future__ import annotations

import logging
import re
from collections.abc import Iterable

_SENSITIVE_PATTERNS = (
    re.compile(r"(?:[A-Za-z]:)?[^\s:*?\"<>|]+\.(?:xlsx|xlsm|xls|xlsb)", re.IGNORECASE),
)


def redact_message(message: str, extra_values: Iterable[str] = ()) -> str:
    """Redact workbook paths and explicitly supplied sensitive values."""
    result = message
    for value in extra_values:
        if value:
            result = result.replace(value, "[redacted]")
    return _SENSITIVE_PATTERNS[0].sub("[workbook]", result)


class RedactingFilter(logging.Filter):
    """Filter that redacts common workbook paths before a record is emitted."""

    def filter(self, record: logging.LogRecord) -> bool:
        record.msg = redact_message(str(record.msg))
        record.args = ()
        return True


def configure_logging(level: str = "INFO") -> logging.Logger:
    logger = logging.getLogger("ex_transform")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    if not logger.handlers:
        handler = logging.StreamHandler()
        handler.addFilter(RedactingFilter())
        handler.setFormatter(logging.Formatter("%(levelname)s %(name)s: %(message)s"))
        logger.addHandler(handler)
    return logger
