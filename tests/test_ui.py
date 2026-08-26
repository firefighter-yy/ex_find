from pathlib import Path

from ex_transform.search_engine import MatchMode, SearchOperator
from ex_transform.ui import SearchViewModel


def test_view_model_requires_prepared_files_and_keywords():
    model = SearchViewModel()
    assert not model.can_search
    model.files.append(Path("book.xlsx"))
    model.add_keywords("alpha")
    assert not model.can_search
    model.stale = False
    assert model.can_search


def test_view_model_keywords_are_trimmed_deduplicated_and_removable():
    model = SearchViewModel(stale=False)
    assert model.add_keywords("  Alpha ")
    assert not model.add_keywords("alpha")
    assert model.keywords == ["Alpha"]
    model.remove_keyword("Alpha")
    assert model.keywords == []


def test_view_model_exposes_search_options_and_export_state():
    model = SearchViewModel(stale=False, files=[Path("book.xlsx")], keywords=["a"])
    model.operator = SearchOperator.OR
    model.match_mode = MatchMode.EXACT
    assert model.can_search
    model.result_count = 2
    assert model.can_export
    model.searching = True
    assert not model.can_export
