"""Pure-Python search over the session-scoped Excel index."""

from __future__ import annotations

import re
import threading
import unicodedata
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from datetime import date, datetime
from datetime import time as datetime_time
from decimal import Decimal, InvalidOperation
from enum import Enum
from pathlib import Path
from typing import Any

from .session_index import IndexedCell, SessionIndex, _cell_from_row


class SearchError(ValueError):
    """Raised for invalid search options."""


class SearchOperator(str, Enum):
    AND = "AND"
    OR = "OR"


class MatchMode(str, Enum):
    CONTAINS = "contains"
    EXACT = "exact"


class FormulaMode(str, Enum):
    RESULT = "result"
    TEXT = "text"
    BOTH = "both"


@dataclass(frozen=True, slots=True)
class SearchResult:
    """One physical worksheet row matching a query."""
    file_id: int
    file_path: Path
    file_order: int
    worksheet_id: int
    worksheet_name: str
    sheet_index: int
    row: int
    cells: tuple[IndexedCell, ...]
    matched_keywords: tuple[str, ...]
    matched_columns: tuple[int, ...]

    @property
    def row_data(self) -> tuple[IndexedCell, ...]:
        return self.cells


@dataclass(frozen=True, slots=True)
class SearchPage:
    results: tuple[SearchResult, ...]
    total_count: int
    offset: int
    limit: int | None
    cancelled: bool = False

    @property
    def has_more(self) -> bool:
        return self.limit is not None and self.offset + len(self.results) < self.total_count


CancelLike = threading.Event | Callable[[], bool] | None
_INVISIBLE = re.compile(r"[\u0000-\u001f\u007f\u00ad\u200b-\u200f\u202a-\u202e\u2060\u2061\u2066-\u2069\ufeff]")
_WHITESPACE = re.compile(r"\s+")


def normalize_display_value(value: Any) -> str:
    """Return the deterministic V1 searchable representation of a value."""
    if value is None:
        return ""
    if isinstance(value, bool):
        text = "TRUE" if value else "FALSE"
    elif isinstance(value, datetime):
        text = value.isoformat(sep=" ")
    elif isinstance(value, (date, datetime_time)):
        text = value.isoformat()
    elif isinstance(value, (float, Decimal)):
        text = _normalize_number(value)
    elif isinstance(value, bytes):
        text = value.hex()
    else:
        text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = _INVISIBLE.sub("", text)
    return _WHITESPACE.sub(" ", text).strip().casefold()


def _normalize_number(value: float | Decimal) -> str:
    try:
        decimal = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return str(value)
    if decimal == 0:
        return "0"
    text = format(decimal, "f").rstrip("0").rstrip(".")
    return text or "0"


def _cancelled(cancel_event: CancelLike) -> bool:
    if cancel_event is None:
        return False
    return cancel_event.is_set() if hasattr(cancel_event, "is_set") else bool(cancel_event())


class SearchService:
    """Search ready files in a :class:`SessionIndex` without Excel or Qt."""

    def __init__(self, index: SessionIndex) -> None:
        self.index = index

    def search(self, keywords: str | Iterable[str], *, operator: SearchOperator | str = SearchOperator.AND,
               match: MatchMode | str = MatchMode.CONTAINS, exact: bool | None = None,
               formula_mode: FormulaMode | str = FormulaMode.RESULT, offset: int = 0,
               limit: int | None = None, cancel_event: CancelLike = None) -> SearchPage:
        terms = _keywords(keywords)
        if exact is not None:
            match = MatchMode.EXACT if exact else MatchMode.CONTAINS
        try:
            query_operator = operator if isinstance(operator, SearchOperator) else SearchOperator(str(operator).upper())
            match_mode = match if isinstance(match, MatchMode) else MatchMode(str(match).lower())
            formula_mode_value = formula_mode if isinstance(formula_mode, FormulaMode) else FormulaMode(str(formula_mode).lower())
        except ValueError as exc:
            raise SearchError("invalid operator, match mode, or formula mode") from exc
        if offset < 0 or (limit is not None and limit < 0):
            raise SearchError("offset and limit must be non-negative")
        if not terms:
            return SearchPage((), 0, offset, limit)
        results: list[SearchResult] = []
        total = 0
        was_cancelled = False
        for group in self._row_groups():
            if _cancelled(cancel_event):
                was_cancelled = True
                break
            result = _match_row(group, terms, query_operator, match_mode, formula_mode_value)
            if result is None:
                continue
            if total >= offset and (limit is None or len(results) < limit):
                results.append(result)
            total += 1
        return SearchPage(tuple(results), total, offset, limit, was_cancelled)

    query = search

    def _row_groups(self):
        rows = self.index.connection.execute(
            """SELECT c.*, f.path AS file_path, f.file_order, w.name AS worksheet_name, w.sheet_index
               FROM cells c JOIN files f ON f.id = c.file_id AND f.status = 'ready'
               JOIN worksheets w ON w.id = c.worksheet_id
               ORDER BY f.file_order, f.id, w.sheet_index, w.id, c.row_number, c.column_number"""
        )
        current_key = None
        cells: list[IndexedCell] = []
        metadata = None
        for row in rows:
            key = (row["file_id"], row["worksheet_id"], row["row_number"])
            if current_key is not None and key != current_key:
                yield metadata, tuple(cells)
                cells = []
            if key != current_key:
                current_key = key
                metadata = row
            cells.append(_cell_from_row(row))
        if current_key is not None:
            yield metadata, tuple(cells)


def _keywords(keywords: str | Iterable[str]) -> tuple[tuple[str, str], ...]:
    source = (keywords,) if isinstance(keywords, str) else keywords
    result: list[tuple[str, str]] = []
    seen: set[str] = set()
    for keyword in source:
        display = str(keyword).strip()
        normalized = normalize_display_value(display)
        if normalized and normalized not in seen:
            result.append((normalized, display))
            seen.add(normalized)
    return tuple(result)


def _matches(value: str, keyword: str, mode: MatchMode) -> bool:
    return value == keyword if mode is MatchMode.EXACT else keyword in value


def _match_row(group, terms, operator, mode, formula_mode):
    metadata, cells = group
    matched_keywords: list[str] = []
    matched_columns: set[int] = set()
    for normalized, display in terms:
        columns: list[int] = []
        for cell in cells:
            candidates = []
            if formula_mode in {FormulaMode.RESULT, FormulaMode.BOTH}:
                candidates.append(normalize_display_value(cell.value))
            if cell.is_formula and formula_mode in {FormulaMode.TEXT, FormulaMode.BOTH}:
                candidates.append(normalize_display_value(cell.formula))
            if any(_matches(candidate, normalized, mode) for candidate in candidates):
                columns.append(cell.column)
        if columns:
            matched_keywords.append(display)
            matched_columns.update(columns)
        elif operator is SearchOperator.AND:
            return None
    if operator is SearchOperator.OR and not matched_keywords:
        return None
    return SearchResult(int(metadata["file_id"]), Path(metadata["file_path"]), int(metadata["file_order"]),
                        int(metadata["worksheet_id"]), metadata["worksheet_name"], int(metadata["sheet_index"]),
                        int(metadata["row_number"]), cells, tuple(matched_keywords), tuple(sorted(matched_columns)))


SearchEngine = SearchService
normalize_value = normalize_display_value




