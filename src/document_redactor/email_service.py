"""이메일(.msg/.eml) → Markdown 렌더·키워드 검색·제거·재검증.

형식 의존성은 어댑터(_load_msg/_load_eml)에만 두고, 나머지 로직은 정규화 구조
EmailContent 위에서 동작해 이메일 파일 없이 단위 테스트할 수 있다. 원본 이메일은
편집하지 않고, 정리 결과는 항상 별도의 <stem>_redacted.md로 산출한다.
"""

from __future__ import annotations

import re as _re
from dataclasses import dataclass, field
from email import message_from_bytes
from email.policy import default as _email_policy
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


_TAG = _re.compile(r"<[^>]+>")


def _strip_html(html: str) -> str:
    """HTML 본문을 대략적인 평문으로 변환한다(태그 제거·공백 정리). 키워드 검색용."""
    text = _TAG.sub(" ", html)
    return _re.sub(r"[ \t]+", " ", text).strip()


def _load_eml(path: Path) -> EmailContent:
    """표준 email 모듈로 .eml을 파싱해 EmailContent로 정규화한다."""
    msg = message_from_bytes(path.read_bytes(), policy=_email_policy)
    body_part = msg.get_body(preferencelist=("plain",))
    if body_part is not None:
        body = body_part.get_content()
    else:
        html_part = msg.get_body(preferencelist=("html",))
        body = _strip_html(html_part.get_content()) if html_part is not None else ""
    attachments = [a.get_filename() or "" for a in msg.iter_attachments()]
    return EmailContent(
        subject=str(msg["subject"] or ""),
        sender=str(msg["from"] or ""),
        to=str(msg["to"] or ""),
        cc=str(msg["cc"] or ""),
        date=str(msg["date"] or ""),
        body=str(body).strip("\n"),
        attachments=[a for a in attachments if a],
    )


def _load_msg(path: Path) -> EmailContent:
    """extract_msg로 .msg를 파싱해 EmailContent로 정규화한다."""
    import extract_msg  # 무거운 의존성 — 함수 내 지연 import

    msg = extract_msg.openMsg(str(path))
    try:
        attachments = [(att.getFilename() or "") for att in msg.attachments]
        return EmailContent(
            subject=str(msg.subject or ""),
            sender=str(msg.sender or ""),
            to=str(msg.to or ""),
            cc=str(msg.cc or ""),
            date=str(msg.date or ""),
            body=str(msg.body or ""),
            attachments=[a for a in attachments if a],
        )
    finally:
        msg.close()


def _load(path: Path) -> EmailContent:
    """확장자로 어댑터를 선택한다(.eml → _load_eml, .msg → _load_msg)."""
    suffix = path.suffix.lower()
    if suffix == ".eml":
        return _load_eml(path)
    if suffix == ".msg":
        return _load_msg(path)
    raise ValueError(f"이메일 형식이 아닙니다: {suffix or '(확장자 없음)'}")
