"""Local Excel information search application."""

__version__ = "0.1.0"

from .excel_ingestion import (
    CellData,
    CellRecord,
    ExcelComReader,
    ExcelIngestion,
    ExcelReader,
    FileReadResult,
    IngestionCancelledError,
    IngestionProgress,
    IngestionResult,
    WorksheetChunk,
    WorksheetInfo,
)

__all__ = [
    "__version__",
    "CellData",
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
]
