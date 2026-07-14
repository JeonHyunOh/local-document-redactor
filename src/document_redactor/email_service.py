"""이메일(.msg/.eml) → Markdown 렌더·키워드 검색·제거·재검증.

형식 의존성은 어댑터(_load_msg/_load_eml)에만 두고, 나머지 로직은 정규화 구조
EmailContent 위에서 동작해 이메일 파일 없이 단위 테스트할 수 있다. 원본 이메일은
편집하지 않고, 정리 결과는 항상 별도의 <stem>_redacted.md로 산출한다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from . import keyword_matcher
from .models import (
    EmailMatch,
    SearchCriteria,
    SearchMode,
)

_BODY_PLACEHOLDER = "(본문 텍스트 없음)"
_RENDER_NOTE = "이메일은 서식·이미지·첨부 내용이 보존되지 않는 평문 .md로 정리됩니다."


@dataclass(frozen=True)
class EmailContent:
    """형식 무관 정규화 이메일 구조. 어댑터가 채우고 렌더·검색이 소비한다."""

    subject: str = ""
    sender: str = ""
    to: str = ""
    cc: str = ""
    date: str = ""
    body: str = ""
    attachments: list[str] = field(default_factory=list)


def _labeled_lines(content: EmailContent) -> list[tuple[str, str]]:
    """(field 라벨, 줄 텍스트) 목록. render_markdown 출력과 줄 순서가 일치한다."""
    lines: list[tuple[str, str]] = [("제목", f"# {content.subject}"), ("", "")]
    if content.sender:
        lines.append(("보낸사람", f"- 보낸사람: {content.sender}"))
    if content.to:
        lines.append(("받는사람", f"- 받는사람: {content.to}"))
    if content.cc:
        lines.append(("참조", f"- 참조: {content.cc}"))
    if content.date:
        lines.append(("날짜", f"- 날짜: {content.date}"))
    if content.attachments:
        lines.append(("첨부", f"- 첨부파일: {', '.join(content.attachments)}"))
    lines.extend([("", ""), ("", "---"), ("", "")])
    body = content.body if content.body else _BODY_PLACEHOLDER
    for body_line in body.split("\n"):
        lines.append(("본문", body_line))
    return lines


def render_markdown(content: EmailContent) -> str:
    """EmailContent를 결정적 Markdown으로 렌더링한다."""
    return "\n".join(text for _, text in _labeled_lines(content))


def _count(line: str, keyword: str, mode: SearchMode, case_sensitive: bool) -> int:
    """한 줄에서 keyword 발견 횟수. EXACT는 줄 전체 일치 시 1, CONTAINS는 부분문자열 개수."""
    if not keyword:
        return 0
    if mode is SearchMode.EXACT:
        return 1 if keyword_matcher.matches(line, keyword, mode, case_sensitive) else 0
    haystack = line if case_sensitive else line.casefold()
    needle = keyword if case_sensitive else keyword.casefold()
    return haystack.count(needle)


def _find_matches(
    content: EmailContent, criteria: SearchCriteria, file_name: str
) -> list[EmailMatch]:
    """렌더된 줄 단위로 키워드 매치를 찾아 EmailMatch 목록으로 반환한다."""
    keywords = keyword_matcher.normalize_keywords(criteria.keywords)
    out: list[EmailMatch] = []
    for index, (label, text) in enumerate(_labeled_lines(content), start=1):
        for keyword in keywords:
            n = _count(text, keyword, criteria.mode, criteria.case_sensitive)
            if n:
                out.append(
                    EmailMatch(
                        file_name=file_name,
                        field=label or "본문",
                        line=index,
                        keyword=keyword,
                        count=n,
                        context=text,
                    )
                )
    return out
