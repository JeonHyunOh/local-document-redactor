"""키워드 정규화와 문자열 매칭.

파일 서비스에 의존하지 않는 순수 로직 계층이다. 검색 단계는 파일을 수정하지 않으며,
이 모듈은 문자열 대상으로만 동작한다. 새 매칭 모드(예: 정규식)는 ``find_matches``의
분기 한 곳만 확장하면 되도록 설계했다.
"""

from __future__ import annotations

from .models import SearchCriteria, SearchMode


def normalize_keywords(keywords: list[str]) -> list[str]:
    """빈 문자열·공백 전용·중복 키워드를 제거하고 입력 순서를 보존한다.

    대소문자는 여기서 접지 않는다(대소문자 구분은 매칭 시점 정책이므로).
    """
    seen: set[str] = set()
    cleaned: list[str] = []
    for raw in keywords:
        keyword = raw.strip()
        if not keyword or keyword in seen:
            continue
        seen.add(keyword)
        cleaned.append(keyword)
    return cleaned


def matches(text: str, keyword: str, mode: SearchMode, case_sensitive: bool) -> bool:
    """단일 텍스트가 단일 키워드에 매칭되는지 판정한다.

    대소문자 구분은 영문에만 의미가 있으며, ``case_sensitive=False``이면 양쪽을
    ``casefold``로 접어 비교한다(한글에는 영향 없음).
    """
    if not keyword:
        return False

    haystack = text if case_sensitive else text.casefold()
    needle = keyword if case_sensitive else keyword.casefold()

    if mode is SearchMode.EXACT:
        return haystack == needle
    if mode is SearchMode.CONTAINS:
        return needle in haystack
    raise ValueError(f"지원하지 않는 검색 모드: {mode!r}")


def find_matches(text: str, criteria: SearchCriteria) -> list[str]:
    """텍스트에 매칭되는 키워드 목록을 반환한다(중복 없이, criteria 순서 유지).

    criteria.keywords는 모델에서 이미 정규화되지만, 직접 호출도 안전하도록
    한 번 더 정규화한다.
    """
    found: list[str] = []
    for keyword in normalize_keywords(criteria.keywords):
        if matches(text, keyword, criteria.mode, criteria.case_sensitive):
            found.append(keyword)
    return found


def remove_keywords(text: str, keywords: list[str], case_sensitive: bool) -> str:
    """텍스트에서 키워드 문자열만 제거한 결과를 반환한다(ExcelAction.REMOVE_KEYWORD용).

    대소문자 무시 시에도 원본 텍스트의 나머지 부분은 보존하며, 매칭된 구간만
    제거한다. 셀/문단 재배치는 하지 않는다.
    """
    result = text
    for keyword in normalize_keywords(keywords):
        if not keyword:
            continue
        if case_sensitive:
            result = result.replace(keyword, "")
        else:
            result = _remove_case_insensitive(result, keyword)
    return result


def _remove_case_insensitive(text: str, keyword: str) -> str:
    """대소문자를 무시하고 keyword와 일치하는 모든 구간을 제거한다."""
    folded_needle = keyword.casefold()
    needle_len = len(keyword)
    out: list[str] = []
    i = 0
    while i < len(text):
        # casefold는 길이를 바꿀 수 있는 문자가 있으나, 한글/영문 일반 텍스트에서는
        # 문자 단위 길이가 보존된다. 안전하게 원본 슬라이스를 접어 비교한다.
        if text[i : i + needle_len].casefold() == folded_needle:
            i += needle_len
        else:
            out.append(text[i])
            i += 1
    return "".join(out)
