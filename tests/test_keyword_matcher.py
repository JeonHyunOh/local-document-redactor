"""keyword_matcher와 SearchCriteria 정규화에 대한 단위 테스트 (M1)."""

from __future__ import annotations

import pytest

from document_redactor.keyword_matcher import (
    find_matches,
    matches,
    normalize_keywords,
    remove_keywords,
)
from document_redactor.models import SearchCriteria, SearchMode


# --------------------------------------------------------------------------- #
# 정규화: 빈 문자열/중복 제거, 순서 보존
# --------------------------------------------------------------------------- #
def test_normalize_removes_blanks_and_strips():
    assert normalize_keywords(["  대외비 ", "", "   ", "내부검토용"]) == [
        "대외비",
        "내부검토용",
    ]


def test_normalize_removes_duplicates_preserving_order():
    assert normalize_keywords(["a", "b", "a", "c", "b"]) == ["a", "b", "c"]


def test_criteria_normalizes_keywords_on_construction():
    criteria = SearchCriteria(keywords=[" 미확정 ", "미확정", ""])
    assert criteria.keywords == ["미확정"]


# --------------------------------------------------------------------------- #
# 포함 검색 (한글)
# --------------------------------------------------------------------------- #
def test_contains_korean():
    assert matches("이 문서는 대외비 자료입니다", "대외비", SearchMode.CONTAINS, False)
    assert not matches("공개 자료입니다", "대외비", SearchMode.CONTAINS, False)


# --------------------------------------------------------------------------- #
# 영문 대소문자 구분
# --------------------------------------------------------------------------- #
def test_case_insensitive_contains():
    assert matches("Confidential Report", "confidential", SearchMode.CONTAINS, False)


def test_case_sensitive_contains():
    assert not matches("Confidential Report", "confidential", SearchMode.CONTAINS, True)
    assert matches("Confidential Report", "Confidential", SearchMode.CONTAINS, True)


# --------------------------------------------------------------------------- #
# 정확히 일치
# --------------------------------------------------------------------------- #
def test_exact_match():
    assert matches("미확정", "미확정", SearchMode.EXACT, False)
    assert not matches("미확정 상태", "미확정", SearchMode.EXACT, False)


def test_exact_match_case_insensitive():
    assert matches("DRAFT", "draft", SearchMode.EXACT, False)
    assert not matches("DRAFT", "draft", SearchMode.EXACT, True)


# --------------------------------------------------------------------------- #
# 여러 키워드
# --------------------------------------------------------------------------- #
def test_find_multiple_matches_preserves_criteria_order():
    criteria = SearchCriteria(
        keywords=["내부검토용", "대외비"], mode=SearchMode.CONTAINS
    )
    found = find_matches("대외비이며 내부검토용 문서", criteria)
    assert found == ["내부검토용", "대외비"]


def test_find_matches_empty_when_none_present():
    criteria = SearchCriteria(keywords=["없는키워드"])
    assert find_matches("아무 내용", criteria) == []


# --------------------------------------------------------------------------- #
# 키워드 제거 (REMOVE_KEYWORD 지원 로직)
# --------------------------------------------------------------------------- #
def test_remove_keywords_case_sensitive():
    assert remove_keywords("대외비 문서 대외비", ["대외비"], True) == " 문서 "


def test_remove_keywords_case_insensitive():
    assert remove_keywords("Confidential and CONFIDENTIAL", ["confidential"], False) == " and "


def test_remove_keywords_leaves_text_without_match():
    assert remove_keywords("공개 자료", ["대외비"], False) == "공개 자료"


# --------------------------------------------------------------------------- #
# 잘못된 모드 방어
# --------------------------------------------------------------------------- #
def test_empty_keyword_never_matches():
    assert not matches("아무 텍스트", "", SearchMode.CONTAINS, False)


def test_unsupported_mode_raises():
    with pytest.raises(ValueError):
        matches("text", "t", "regex", False)  # type: ignore[arg-type]
