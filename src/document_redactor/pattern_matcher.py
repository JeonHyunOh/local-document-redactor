"""정형 개인정보 패턴(전화·이메일·주민번호·카드·계좌)을 감지·제거하는 순수 로직.

키워드와 독립적으로 항상 적용된다(계좌번호만 opt-in). 파일 I/O 없음. 겹치는 매치는
긴 것을 우선해 비겹침으로 제거한다. 계좌번호는 오탐 위험이 커 기본 비활성이다.
"""

from __future__ import annotations

import re

# 구체적 패턴을 먼저(겹침 시 우선). 유형명은 UI 라벨로 쓰인다.
_ALWAYS: list[tuple[str, re.Pattern]] = [
    ("이메일", re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")),
    ("주민등록번호", re.compile(r"\b\d{6}-\d{7}\b")),
    ("신용카드", re.compile(r"\b\d{4}[- ]\d{4}[- ]\d{4}[- ]\d{4}\b")),
    ("전화번호", re.compile(r"(?:\+?82[- ]?)?\b0\d{1,2}[- ]?\d{3,4}[- ]?\d{4}\b")),
]
_ACCOUNT: tuple[str, re.Pattern] = (
    "계좌번호",
    re.compile(r"\b\d{2,6}-\d{2,6}-\d{2,6}(?:-\d{1,6})?\b"),
)


def _active(include_account: bool) -> list[tuple[str, re.Pattern]]:
    return _ALWAYS + ([_ACCOUNT] if include_account else [])


def _spans(text: str, include_account: bool) -> list[tuple[int, int, str, str]]:
    """(start, end, 유형, 매칭값) 목록. 겹치면 긴 매치 우선으로 비겹침 선택한다."""
    raw: list[tuple[int, int, str, str]] = []
    for label, pattern in _active(include_account):
        for m in pattern.finditer(text):
            raw.append((m.start(), m.end(), label, m.group()))
    raw.sort(key=lambda s: (s[0], -(s[1] - s[0])))  # 시작 오름차순, 길이 내림차순
    chosen: list[tuple[int, int, str, str]] = []
    last_end = -1
    for start, end, label, value in raw:
        if start >= last_end:
            chosen.append((start, end, label, value))
            last_end = end
    return chosen


def find_patterns(text: str, include_account: bool = False) -> list[tuple[str, str]]:
    """텍스트에서 감지된 (유형, 매칭값) 목록을 반환한다."""
    return [(label, value) for _, _, label, value in _spans(text, include_account)]


def remove_patterns(text: str, include_account: bool = False) -> str:
    """감지된 모든 패턴 구간을 제거한 문자열을 반환한다."""
    result = text
    for start, end, _, _ in sorted(_spans(text, include_account), reverse=True):
        result = result[:start] + result[end:]
    return result


def contains_pattern(text: str, include_account: bool = False) -> bool:
    """텍스트에 감지되는 패턴이 하나라도 있으면 True."""
    return bool(_spans(text, include_account))
