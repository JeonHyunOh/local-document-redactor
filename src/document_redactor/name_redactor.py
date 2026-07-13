"""파일명·폴더명에서 키워드를 제거하는 순수 로직.

파일 I/O에 의존하지 않는다. keyword_matcher의 문자열 제거를 재사용하고, 그 위에
이름 특유의 처리(확장자 보존·잔재 정리·충돌 회피)를 얹는다. 이름 검사는 파일 내용
검사와 독립적이며, 폴더명에도 동일 로직을 적용한다.
"""

from __future__ import annotations

import re
from pathlib import PurePosixPath

from . import keyword_matcher

_SEPARATOR_RUN = re.compile(r"[ _-]{2,}")
_TRIM_CHARS = " _-."
_FALLBACK_STEM = "redacted"


def name_contains_keyword(name: str, keywords: list[str], case_sensitive: bool) -> bool:
    """이름(파일명 또는 폴더명)에 등록 키워드가 부분문자열로 존재하는지 판정한다."""
    haystack = name if case_sensitive else name.casefold()
    for keyword in keyword_matcher.normalize_keywords(keywords):
        needle = keyword if case_sensitive else keyword.casefold()
        if needle and needle in haystack:
            return True
    return False


def _split_suffix(name: str) -> tuple[str, str]:
    """(stem, suffix)로 나눈다. 마지막 확장자만 suffix로 취급한다."""
    suffix = PurePosixPath(name).suffix
    stem = name[: len(name) - len(suffix)] if suffix else name
    return stem, suffix


def _tidy(stem: str) -> str:
    """키워드 제거 후 남은 구분자 잔재를 정리한다(연속 구분자 축약·앞뒤 트리밍). 비면 fallback."""
    collapsed = _SEPARATOR_RUN.sub("_", stem)
    trimmed = collapsed.strip(_TRIM_CHARS)
    return trimmed or _FALLBACK_STEM


def redact_segment(name: str, keywords: list[str], case_sensitive: bool) -> str:
    """폴더명 등 확장자 없는 이름에서 키워드를 제거하고 정리한다."""
    removed = keyword_matcher.remove_keywords(name, keywords, case_sensitive)
    return _tidy(removed)


def redact_filename(name: str, keywords: list[str], case_sensitive: bool) -> str:
    """파일명에서 키워드를 제거한다. 마지막 확장자(suffix)는 보존한다."""
    stem, suffix = _split_suffix(name)
    return _tidy(keyword_matcher.remove_keywords(stem, keywords, case_sensitive)) + suffix


def unique_name(candidate: str, taken: set[str]) -> str:
    """taken에 이미 있으면 stem 뒤에 _1, _2 …를 붙여 고유 이름을 만든다(확장자 유지).

    taken은 호출자가 관리하는 '이미 사용 중인 이름' 집합이다. 반환값을 taken에 추가하는
    책임은 호출자에게 있다.
    """
    if candidate not in taken:
        return candidate
    stem, suffix = _split_suffix(candidate)
    index = 1
    while True:
        trial = f"{stem}_{index}{suffix}"
        if trial not in taken:
            return trial
        index += 1
