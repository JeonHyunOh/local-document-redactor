# 이메일(.msg/.eml) → .md 키워드 제거 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `.msg`·`.eml` 이메일을 분석해 제목·헤더·첨부목록·본문을 Markdown으로 렌더링하고, 키워드를 제거한 `<stem>_redacted.md`를 산출한다(원본 이메일 보존).

**Architecture:** 형식 의존성을 얇은 어댑터(`_load_msg`=extract_msg, `_load_eml`=표준 email) 두 개에 격리하고, 정규화 구조 `EmailContent` 위에서 렌더·검색·제거·검증을 공유한다. `excel_service`/`pdf_service`와 같은 서비스 계약(`search`/`apply_edit`/`verify`)을 구현하는 `email_service.py`를 추가하고 `file_service`가 `.msg`/`.eml`/`.md`를 라우팅한다. UI(`app.py`)에는 로직을 넣지 않는다.

**Tech Stack:** Python 3.11+, extract-msg(신규), 표준 `email`, openpyxl, PyMuPDF, pydantic v2, Streamlit, pytest.

## Global Constraints

- 원본 이메일 파일을 절대 편집하지 않는다. 산출물은 별도의 새 `.md` 파일.
- 검색(무수정)과 편집은 별개 단계. 이메일 편집은 **모든 키워드 occurrence 제거**로 고정(부분 선택 미지원).
- 이름에서의 키워드 제거는 **항상 부분문자열**(내용 모드가 EXACT여도). `case_sensitive`는 이름에도 적용.
- 제자리 교체 모드는 이메일 **내용**을 처리하지 않는다(형식이 `.md`로 바뀌므로). 파일명 rename만 적용.
- 예외를 `except Exception: pass`로 삼키지 않는다. 파일별 오류는 `error`에 기록하고 배치는 계속.
- 경로 처리는 `pathlib`만. 서비스 함수는 Streamlit에 의존하지 않는다.
- `email_service`는 `file_service`를 import하지 않는다(순환 방지). `.eml`은 표준 라이브러리만 사용.
- 테스트 픽스처 파일은 실행 시 Python으로 생성(레포에 바이너리 금지).

---

## File Structure

- Create: `src/document_redactor/email_service.py` — 이메일 파싱·렌더·검색·편집·검증.
- Create: `tests/conftest.py` — `make_eml` / `make_msg` 픽스처 팩토리.
- Create: `tests/test_email_service.py` — 순수 로직 + 어댑터 + 서비스 계약 테스트.
- Modify: `src/document_redactor/models.py` — `FileType`(MSG/EML/MD), `EmailMatch`, `SearchReport.email_matches`, `BatchEditItem.note`.
- Modify: `src/document_redactor/file_service.py` — `.msg`/`.eml` 업로드 허용, `.msg`/`.eml`/`.md` 라우팅.
- Modify: `src/document_redactor/batch_service.py` — `scan_folder` 확장자, in-place 이메일 스킵+rename.
- Modify: `tests/test_file_service.py`, `tests/test_batch_service.py` — 라우팅·배치 시나리오.
- Modify: `app.py` — 업로더 확장자, 단일 파일 이메일 미리보기/.md 다운로드.
- Modify: `pyproject.toml` — `extract-msg` 의존성 추가.

---

## Task 1: 데이터 모델 확장

**Files:**
- Modify: `src/document_redactor/models.py`
- Test: `tests/test_email_service.py` (일부는 Task 2에서 사용)

**Interfaces:**
- Produces:
  - `FileType.MSG = "msg"`, `FileType.EML = "eml"`, `FileType.MD = "md"`
  - `EmailMatch(file_name: str, field: str, line: int, keyword: str, count: int, context: str)`
  - `SearchReport.email_matches: list[EmailMatch]` (기본 빈 목록), `total_matches`에 합산
  - `BatchEditItem.note: str | None = None`

- [ ] **Step 1: Write the failing test**

`tests/test_email_service.py`(새 파일) 상단:

```python
"""email_service 및 관련 모델 테스트 — 이메일 파싱·렌더·검색·편집·검증."""

from __future__ import annotations

from pathlib import Path

from document_redactor.models import (
    EmailMatch,
    FileType,
    SearchCriteria,
    SearchReport,
)


def test_filetype_has_email_members():
    assert FileType.MSG.value == "msg"
    assert FileType.EML.value == "eml"
    assert FileType.MD.value == "md"


def test_searchreport_total_includes_email_matches():
    report = SearchReport(
        file_name="a.eml",
        file_type=FileType.EML,
        criteria=SearchCriteria(keywords=["x"]),
        email_matches=[
            EmailMatch(file_name="a.eml", field="본문", line=3, keyword="x", count=2, context="x x"),
        ],
    )
    assert report.total_matches == 2
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv/Scripts/python.exe -m pytest tests/test_email_service.py -q`
Expected: FAIL — `AttributeError: MSG` / `EmailMatch` import 오류.

- [ ] **Step 3: Implement**

`models.py`의 `FileType`에 멤버 추가:

```python
class FileType(str, Enum):
    """지원 입력 유형(XLSX/XLSM/PDF/MSG/EML)과 렌더 출력 유형(MD)."""

    XLSX = "xlsx"
    XLSM = "xlsm"
    PDF = "pdf"
    MSG = "msg"
    EML = "eml"
    MD = "md"  # 이메일을 정리해 렌더한 Markdown 산출물(입력 유형 아님)
```

`PdfMatch` 클래스 정의 아래에 `EmailMatch` 추가:

```python
class EmailMatch(BaseModel):
    """이메일을 렌더한 Markdown의 매치 한 건."""

    file_name: str
    field: str                # "제목"/"보낸사람"/"받는사람"/"참조"/"날짜"/"첨부"/"본문"
    line: int                 # 렌더된 .md의 1-기반 줄 번호
    keyword: str
    count: int                # 해당 줄에서 발견된 횟수
    context: str              # 해당 줄 텍스트
```

`SearchReport`에 필드 추가 + `total_matches` 갱신:

```python
class SearchReport(BaseModel):
    file_name: str
    file_type: FileType
    criteria: SearchCriteria
    excel_matches: list[ExcelMatch] = Field(default_factory=list)
    pdf_matches: list[PdfMatch] = Field(default_factory=list)
    email_matches: list[EmailMatch] = Field(default_factory=list)
    notes: list[str] = Field(default_factory=list)

    @property
    def total_matches(self) -> int:
        return (
            len(self.excel_matches)
            + sum(m.count for m in self.pdf_matches)
            + sum(m.count for m in self.email_matches)
        )
```

`BatchEditItem`에 `note` 필드 추가(기존 필드 유지):

```python
    renamed_to: str | None = None
    note: str | None = None  # 비오류 안내(예: 제자리 모드에서 이메일 내용 미지원)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv/Scripts/python.exe -m pytest tests/test_email_service.py -q`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
git add src/document_redactor/models.py tests/test_email_service.py
git commit -m "feat: 이메일 매치·FileType(MSG/EML/MD)·BatchEditItem.note 모델 추가"
```

---

## Task 2: `email_service` 순수 로직 (EmailContent·렌더·검색)

**Files:**
- Create: `src/document_redactor/email_service.py`
- Test: `tests/test_email_service.py`

**Interfaces:**
- Consumes: `keyword_matcher`(normalize_keywords, matches), `models`(EmailMatch, SearchMode).
- Produces:
  - `EmailContent`(dataclass: subject, sender, to, cc, date, body, attachments: list[str])
  - `render_markdown(content: EmailContent) -> str`
  - `_labeled_lines(content: EmailContent) -> list[tuple[str, str]]`
  - `_find_matches(content: EmailContent, criteria: SearchCriteria, file_name: str) -> list[EmailMatch]`

- [ ] **Step 1: Write the failing test**

`tests/test_email_service.py`에 추가:

```python
from document_redactor import email_service
from document_redactor.email_service import EmailContent


def _content(**kw):
    base = dict(subject="", sender="", to="", cc="", date="", body="", attachments=[])
    base.update(kw)
    return EmailContent(**base)


def test_render_markdown_omits_empty_fields():
    md = email_service.render_markdown(_content(subject="대외비 건", sender="홍길동"))
    assert "# 대외비 건" in md
    assert "- 보낸사람: 홍길동" in md
    assert "받는사람" not in md   # to 비어있으면 줄 생략


def test_render_markdown_lists_attachments_and_body():
    md = email_service.render_markdown(
        _content(subject="s", body="본문 내용", attachments=["a.xlsx", "b.pdf"])
    )
    assert "- 첨부파일: a.xlsx, b.pdf" in md
    assert "본문 내용" in md


def test_render_markdown_empty_body_placeholder():
    md = email_service.render_markdown(_content(subject="s"))
    assert "(본문 텍스트 없음)" in md


def test_find_matches_locates_keyword_in_body_with_field_and_count():
    content = _content(subject="공개", body="포스코 포스코 관련")
    matches = email_service._find_matches(content, SearchCriteria(keywords=["포스코"]), "a.eml")
    assert len(matches) == 1
    m = matches[0]
    assert m.field == "본문" and m.count == 2 and "포스코" in m.context


def test_find_matches_in_subject_and_attachment():
    content = _content(subject="포스코 보고", attachments=["포스코_첨부.xlsx"])
    kws = SearchCriteria(keywords=["포스코"])
    fields = {m.field for m in email_service._find_matches(content, kws, "a.eml")}
    assert "제목" in fields and "첨부" in fields


def test_find_matches_case_insensitive_default_and_sensitive():
    content = _content(subject="POSCO note")
    assert email_service._find_matches(content, SearchCriteria(keywords=["posco"]), "a") 
    assert not email_service._find_matches(
        content, SearchCriteria(keywords=["posco"], case_sensitive=True), "a"
    )
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv/Scripts/python.exe -m pytest tests/test_email_service.py -q -k "render or find_matches"`
Expected: FAIL — `ModuleNotFoundError: email_service` / 함수 미정의.

- [ ] **Step 3: Implement**

`src/document_redactor/email_service.py`:

```python
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
    EditRequest,
    EditResult,
    EmailMatch,
    FileType,
    SearchCriteria,
    SearchMode,
    SearchReport,
    VerificationResult,
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv/Scripts/python.exe -m pytest tests/test_email_service.py -q -k "render or find_matches"`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/document_redactor/email_service.py tests/test_email_service.py
git commit -m "feat: email_service 순수 로직(EmailContent·render_markdown·검색)"
```

---

## Task 3: 형식 어댑터 (`_load_eml`/`_load_msg`) + 픽스처 + 의존성

**Files:**
- Modify: `src/document_redactor/email_service.py`
- Modify: `pyproject.toml`
- Create: `tests/conftest.py`
- Test: `tests/test_email_service.py`

**Interfaces:**
- Consumes: `extract_msg`(신규), 표준 `email`.
- Produces:
  - `_load_eml(path: Path) -> EmailContent`
  - `_load_msg(path: Path) -> EmailContent`
  - `_load(path: Path) -> EmailContent` (확장자로 분기)

- [ ] **Step 1: pyproject에 extract-msg 추가 + 설치**

`pyproject.toml`의 `dependencies`에 추가:

```toml
dependencies = [
    "streamlit>=1.33",
    "openpyxl>=3.1",
    "PyMuPDF>=1.24",
    "pydantic>=2.6",
    "extract-msg>=0.54",
]
```

Run: `.venv/Scripts/python.exe -m pip install -e ".[dev]"`
Expected: `extract-msg`(및 의존성) 설치 완료.

- [ ] **Step 2: Write the failing test (픽스처 + 어댑터)**

`tests/conftest.py`(새 파일):

```python
"""테스트용 이메일 픽스처 팩토리(.eml=표준 email, .msg=extract_msg OleWriter)."""

from __future__ import annotations

from pathlib import Path

import pytest


@pytest.fixture
def make_eml():
    from email.message import EmailMessage

    def _make(path: Path, *, subject="", body="", sender="", to="", cc="", attachments=()):
        path.parent.mkdir(parents=True, exist_ok=True)
        msg = EmailMessage()
        if subject:
            msg["Subject"] = subject
        if sender:
            msg["From"] = sender
        if to:
            msg["To"] = to
        if cc:
            msg["Cc"] = cc
        msg.set_content(body)
        for name in attachments:
            msg.add_attachment(
                b"x", maintype="application", subtype="octet-stream", filename=name
            )
        path.write_bytes(msg.as_bytes())
        return path

    return _make


@pytest.fixture
def make_msg():
    from extract_msg import OleWriter

    def _make(path: Path, *, subject="", body="", sender=""):
        path.parent.mkdir(parents=True, exist_ok=True)
        writer = OleWriter()

        def prop(propid: int, text: str) -> None:
            if text:
                writer.addEntry(f"__substg1.0_{propid:04X}001F", text.encode("utf-16-le"))

        prop(0x001A, "IPM.Note")  # PidTagMessageClass (유효 MSG 판별에 필요)
        prop(0x0037, subject)     # PidTagSubject
        prop(0x1000, body)        # PidTagBody
        prop(0x0C1A, sender)      # PidTagSenderName
        writer.addEntry("__properties_version1.0", b"\x00" * 32)
        writer.write(str(path))
        return path

    return _make
```

`tests/test_email_service.py`에 추가:

```python
def test_load_eml_normalizes_all_fields(make_eml, tmp_path):
    p = make_eml(
        tmp_path / "m.eml",
        subject="대외비 건",
        body="포스코 본문",
        sender="홍길동 <hong@example.com>",
        to="김철수 <kim@example.com>",
        cc="이영희 <lee@example.com>",
        attachments=["포스코_첨부.xlsx"],
    )
    content = email_service._load(p)
    assert content.subject == "대외비 건"
    assert "홍길동" in content.sender
    assert "김철수" in content.to
    assert "이영희" in content.cc
    assert "포스코 본문" in content.body
    assert content.attachments == ["포스코_첨부.xlsx"]


def test_load_msg_normalizes_core_fields(make_msg, tmp_path):
    p = make_msg(tmp_path / "m.msg", subject="대외비 건", body="포스코 본문", sender="홍길동")
    content = email_service._load(p)
    assert content.subject == "대외비 건"
    assert content.sender == "홍길동"
    assert "포스코 본문" in content.body
```

- [ ] **Step 3: Run test to verify it fails**

Run: `.venv/Scripts/python.exe -m pytest tests/test_email_service.py -q -k "load_eml or load_msg"`
Expected: FAIL — `_load` 미정의.

- [ ] **Step 4: Implement**

`email_service.py` 상단 import에 추가:

```python
from email import message_from_bytes
from email.policy import default as _email_policy
```

파일 하단에 어댑터 추가:

```python
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
```

그리고 파일 하단에 HTML 평문화 헬퍼 추가:

```python
import re as _re

_TAG = _re.compile(r"<[^>]+>")


def _strip_html(html: str) -> str:
    """HTML 본문을 대략적인 평문으로 변환한다(태그 제거·공백 정리). 키워드 검색용."""
    text = _TAG.sub(" ", html)
    return _re.sub(r"[ \t]+", " ", text).strip()
```

- [ ] **Step 5: Run test to verify it passes**

Run: `.venv/Scripts/python.exe -m pytest tests/test_email_service.py -q -k "load_eml or load_msg"`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add pyproject.toml tests/conftest.py src/document_redactor/email_service.py tests/test_email_service.py
git commit -m "feat: 이메일 어댑터(.eml/.msg) + extract-msg 의존성 + 픽스처"
```

---

## Task 4: 서비스 계약 (`search`/`apply_edit`/`verify`)

**Files:**
- Modify: `src/document_redactor/email_service.py`
- Test: `tests/test_email_service.py`

**Interfaces:**
- Consumes: Task 2/3 (`_load`, `render_markdown`, `_find_matches`), `keyword_matcher.remove_keywords`, `models`.
- Produces:
  - `search(path: Path, criteria: SearchCriteria) -> SearchReport`
  - `apply_edit(path: Path, request: EditRequest, output_dir: Path, selected=None) -> EditResult`
  - `verify(output_path: Path, criteria: SearchCriteria) -> VerificationResult`

- [ ] **Step 1: Write the failing test**

`tests/test_email_service.py`에 추가:

```python
from document_redactor.models import EditRequest


def test_search_reports_email_matches(make_eml, tmp_path):
    p = make_eml(tmp_path / "m.eml", subject="포스코 보고", body="대외비 내용")
    report = email_service.search(p, SearchCriteria(keywords=["포스코", "대외비"]))
    assert report.file_type == FileType.EML
    assert report.total_matches == 2
    assert {m.keyword for m in report.email_matches} == {"포스코", "대외비"}


def test_apply_edit_produces_redacted_md_and_keeps_original(make_eml, tmp_path):
    p = make_eml(tmp_path / "m.eml", subject="포스코 보고", body="포스코 대외비 내용")
    out = tmp_path / "out"
    req = EditRequest(criteria=SearchCriteria(keywords=["포스코"]))
    result = email_service.apply_edit(p, req, out)

    produced = Path(result.output_path)
    assert produced.name == "m_redacted.md"
    text = produced.read_text(encoding="utf-8")
    assert "포스코" not in text
    assert "대외비 내용" in text          # 다른 텍스트는 보존
    assert result.file_type == FileType.MD
    assert result.redactions_applied == 2  # 제목 1 + 본문 1
    # 원본 .eml 미수정
    assert "포스코" in email_service._load(p).subject


def test_verify_clean_after_edit(make_eml, tmp_path):
    p = make_eml(tmp_path / "m.eml", subject="포스코", body="포스코")
    out = tmp_path / "out"
    req = EditRequest(criteria=SearchCriteria(keywords=["포스코"]))
    result = email_service.apply_edit(p, req, out)
    verification = email_service.verify(Path(result.output_path), req.criteria)
    assert verification.clean is True


def test_verify_detects_remaining(tmp_path):
    md = tmp_path / "x_redacted.md"
    md.write_text("# 포스코 남음\n\n본문", encoding="utf-8")
    verification = email_service.verify(md, SearchCriteria(keywords=["포스코"]))
    assert verification.clean is False
    assert verification.remaining is not None
    assert verification.remaining.total_matches == 1


def test_roundtrip_msg(make_msg, tmp_path):
    p = make_msg(tmp_path / "m.msg", subject="포스코 건", body="대외비 본문")
    out = tmp_path / "out"
    req = EditRequest(criteria=SearchCriteria(keywords=["포스코", "대외비"]))
    result = email_service.apply_edit(p, req, out)
    assert Path(result.output_path).name == "m_redacted.md"
    verification = email_service.verify(Path(result.output_path), req.criteria)
    assert verification.clean is True
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv/Scripts/python.exe -m pytest tests/test_email_service.py -q -k "search_reports or apply_edit or verify or roundtrip"`
Expected: FAIL — `search`/`apply_edit`/`verify` 미정의.

- [ ] **Step 3: Implement**

`email_service.py` 파일 하단에 추가:

```python
def _occurrences(text: str, keywords: list[str], case_sensitive: bool) -> int:
    """전체 텍스트에서 키워드(부분문자열) 총 발견 횟수. 제거 건수 집계용."""
    total = 0
    haystack = text if case_sensitive else text.casefold()
    for keyword in keyword_matcher.normalize_keywords(keywords):
        needle = keyword if case_sensitive else keyword.casefold()
        if needle:
            total += haystack.count(needle)
    return total


def search(path: Path, criteria: SearchCriteria) -> SearchReport:
    """이메일을 파싱·렌더해 키워드를 검색한다(파일 무수정)."""
    path = Path(path)
    content = _load(path)
    file_type = FileType.MSG if path.suffix.lower() == ".msg" else FileType.EML
    return SearchReport(
        file_name=path.name,
        file_type=file_type,
        criteria=criteria,
        email_matches=_find_matches(content, criteria, path.name),
        notes=[_RENDER_NOTE],
    )


def apply_edit(
    path: Path, request: EditRequest, output_dir: Path, selected=None
) -> EditResult:
    """이메일을 렌더한 뒤 키워드를 제거해 <stem>_redacted.md로 저장한다.

    selected는 무시한다(이메일은 항상 전체 키워드 제거). 원본 이메일은 편집하지 않는다.
    """
    path = Path(path)
    content = _load(path)
    text = render_markdown(content)
    keywords = request.criteria.keywords
    case_sensitive = request.criteria.case_sensitive
    redacted = keyword_matcher.remove_keywords(text, keywords, case_sensitive)
    removed = _occurrences(text, keywords, case_sensitive)

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / f"{path.stem}_redacted.md"
    out_path.write_text(redacted, encoding="utf-8")

    return EditResult(
        source_name=path.name,
        output_path=str(out_path),
        file_type=FileType.MD,
        redactions_applied=removed,
        log=[f"이메일에서 키워드 {removed}건 제거 → {out_path.name}"],
    )


def verify(output_path: Path, criteria: SearchCriteria) -> VerificationResult:
    """산출된 .md 텍스트를 재검색해 키워드 잔존 여부를 확인한다."""
    output_path = Path(output_path)
    text = output_path.read_text(encoding="utf-8")
    matches: list[EmailMatch] = []
    keywords = keyword_matcher.normalize_keywords(criteria.keywords)
    for index, line in enumerate(text.split("\n"), start=1):
        for keyword in keywords:
            n = _count(line, keyword, criteria.mode, criteria.case_sensitive)
            if n:
                matches.append(
                    EmailMatch(
                        file_name=output_path.name,
                        field="본문",
                        line=index,
                        keyword=keyword,
                        count=n,
                        context=line,
                    )
                )
    remaining = (
        SearchReport(
            file_name=output_path.name,
            file_type=FileType.MD,
            criteria=criteria,
            email_matches=matches,
        )
        if matches
        else None
    )
    return VerificationResult(output_path=str(output_path), clean=not matches, remaining=remaining)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv/Scripts/python.exe -m pytest tests/test_email_service.py -q`
Expected: PASS (전체 email_service 테스트 그린)

- [ ] **Step 5: Commit**

```bash
git add src/document_redactor/email_service.py tests/test_email_service.py
git commit -m "feat: email_service search/apply_edit/verify 계약 구현"
```

---

## Task 5: `file_service` 라우팅

**Files:**
- Modify: `src/document_redactor/file_service.py`
- Test: `tests/test_file_service.py`

**Interfaces:**
- Consumes: `email_service`(search/apply_edit/verify), `FileType.MSG/EML`.
- Produces: `.msg`/`.eml` 업로드 허용, `_service_for`가 `.msg`/`.eml`/`.md` → email_service.

- [ ] **Step 1: Write the failing test**

`tests/test_file_service.py` 하단에 추가(파일 상단 import에 `email_service`, `pytest` 있는지 확인, 없으면 추가):

```python
from document_redactor import email_service  # 상단 import 블록에 추가


def test_detect_file_type_accepts_email():
    assert file_service.detect_file_type("a.msg") is FileType.MSG
    assert file_service.detect_file_type("a.eml") is FileType.EML


def test_detect_file_type_rejects_md_input():
    import pytest

    with pytest.raises(file_service.UnsupportedFileError):
        file_service.detect_file_type("a.md")


def test_service_for_routes_email_and_md():
    assert file_service._service_for(Path("a.msg")) is email_service
    assert file_service._service_for(Path("a.eml")) is email_service
    assert file_service._service_for(Path("out_redacted.md")) is email_service
```

(파일 상단에 `from document_redactor.models import FileType`, `from pathlib import Path`가 없으면 추가.)

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv/Scripts/python.exe -m pytest tests/test_file_service.py -q -k "email or md"`
Expected: FAIL — `.msg` 미지원 / `_service_for`가 `.md`에서 예외.

- [ ] **Step 3: Implement**

`file_service.py` 상단 import에 `email_service` 추가:

```python
from . import email_service, excel_service, pdf_service
```

`_EXTENSION_MAP`에 이메일 추가:

```python
_EXTENSION_MAP: dict[str, FileType] = {
    ".xlsx": FileType.XLSX,
    ".xlsm": FileType.XLSM,
    ".pdf": FileType.PDF,
    ".msg": FileType.MSG,
    ".eml": FileType.EML,
}
```

`detect_file_type`의 안내 문구를 갱신:

```python
        raise UnsupportedFileError(
            f"지원하지 않는 파일 형식입니다: {suffix or '(확장자 없음)'}. "
            "지원 형식은 .xlsx, .xlsm, .pdf, .msg, .eml 입니다. "
            "(.xls, 암호화 파일, 스캔 전용 PDF는 지원하지 않습니다.)"
        )
```

`_service_for`를 이메일·산출물(.md) 라우팅까지 처리하도록 교체:

```python
def _service_for(path: Path):
    """파일 경로에 맞는 서비스 모듈을 반환한다.

    산출물 재검증을 위해 .md는 email_service로 라우팅한다(업로드 검증에는 .md를
    포함하지 않으므로 사용자가 .md를 입력할 수는 없다).
    """
    suffix = path.suffix.lower()
    if suffix == ".md":
        return email_service
    file_type = detect_file_type(path.name)
    if file_type in (FileType.MSG, FileType.EML):
        return email_service
    if file_type is FileType.PDF:
        return pdf_service
    return excel_service
```

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv/Scripts/python.exe -m pytest tests/test_file_service.py -q`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/document_redactor/file_service.py tests/test_file_service.py
git commit -m "feat: file_service가 .msg/.eml 업로드 허용 + .md 재검증 라우팅"
```

---

## Task 6: `batch_service` — 스캔 확장자 + in-place 이메일 처리

**Files:**
- Modify: `src/document_redactor/batch_service.py`
- Test: `tests/test_batch_service.py`

**Interfaces:**
- Consumes: `email_service`(라우팅 경유), `FileType`.
- Produces:
  - `scan_folder`가 `.msg`/`.eml` 포함.
  - `batch_edit`: 이메일 내용 매치 → `<stem>_redacted.md` 산출(기존 로직 재사용, 코드 변경 없음).
  - `batch_edit_in_place`: 이메일 내용 미처리(원본 유지 + `note`), 파일명 키워드 시 Phase 2 rename 적용.

- [ ] **Step 1: Write the failing test**

`tests/test_batch_service.py` 하단에 추가(파일에 `make_eml` 픽스처는 conftest에서 자동 주입).
파일 상단 import에 `from document_redactor import email_service`를 추가한다(테스트가 `_load` 사용):

```python
def test_scan_folder_includes_email(tmp_path: Path, make_eml):
    _xlsx(tmp_path / "a.xlsx", "x")
    make_eml(tmp_path / "b.eml", subject="s", body="x")
    names = {p.name for p in batch_service.scan_folder(tmp_path, recursive=True)}
    assert names == {"a.xlsx", "b.eml"}


def test_batch_edit_email_produces_markdown(tmp_path: Path, make_eml):
    root = tmp_path / "src"
    make_eml(root / "메일.eml", subject="포스코 보고", body="대외비 내용")
    out = tmp_path / "out"
    req = EditRequest(criteria=_criteria("포스코", "대외비"), excel_action=ExcelAction.CLEAR_CELL)
    items = batch_service.batch_edit(root, req, out, recursive=True)

    produced = out / "메일_redacted.md"
    assert produced.exists()
    text = produced.read_text(encoding="utf-8")
    assert "포스코" not in text and "대외비" not in text
    assert any(i.output_path for i in items)
    # 원본 보존
    assert (root / "메일.eml").exists()


def test_batch_edit_in_place_skips_email_content_but_renames(tmp_path: Path, make_eml):
    root = tmp_path / "src"
    make_eml(root / "포스코_메일.eml", subject="포스코", body="대외비")
    backup = tmp_path / "src_backup"
    req = EditRequest(criteria=_criteria("포스코", "대외비"), excel_action=ExcelAction.CLEAR_CELL)
    items = batch_service.batch_edit_in_place(root, req, backup, recursive=True)
    item = items[0]

    # 내용은 처리하지 않음 → 원본 .eml 그대로(본문에 대외비 남아 있음)
    renamed = root / "메일.eml"
    assert renamed.exists() and not (root / "포스코_메일.eml").exists()  # 파일명만 rename
    # 원본 내용 미수정 확인은 파싱해서(전송 인코딩 무관) — subject의 '포스코', body의 '대외비' 잔존
    reloaded = email_service._load(renamed)
    assert "포스코" in reloaded.subject and "대외비" in reloaded.body
    assert item.note is not None and "제자리" in item.note
    assert item.renamed_to == "메일.eml"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv/Scripts/python.exe -m pytest tests/test_batch_service.py -q -k "email"`
Expected: FAIL — `scan_folder`가 `.eml` 제외 / in-place가 이메일 내용을 편집하려다 실패하거나 `note` 없음.

- [ ] **Step 3: Implement**

`batch_service.py`의 `_SUPPORTED_SUFFIXES`에 이메일 추가:

```python
_SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".pdf", ".msg", ".eml"}
```

`batch_service.py` 상단 import에 `FileType` 추가(없으면):

```python
from .models import (
    BatchEditItem,
    BatchSearchItem,
    EditRequest,
    FileType,
    SearchCriteria,
)
```

모듈 상단(함수 밖)에 헬퍼 추가:

```python
_EMAIL_SUFFIXES = {".msg", ".eml"}
_INPLACE_EMAIL_NOTE = "제자리 모드는 이메일 내용을 지원하지 않습니다(별도 출력 폴더 모드를 사용하세요)."
```

`batch_edit_in_place`의 **Phase 1 루프** 안, `relative = ...` 다음 줄에 이메일 스킵 분기를 추가한다. 기존:

```python
    for index, path in enumerate(files, start=1):
        relative = path.relative_to(root).as_posix()
        try:
            report = file_service.search(path, request.criteria)
```

를 아래로 교체:

```python
    for index, path in enumerate(files, start=1):
        relative = path.relative_to(root).as_posix()
        try:
            if path.suffix.lower() in _EMAIL_SUFFIXES:
                # 이메일은 제자리 교체(형식 변경) 불가 — 내용 미처리, 파일명 rename은 Phase 2에서.
                item = BatchEditItem(
                    path=str(path), relative_path=relative, note=_INPLACE_EMAIL_NOTE
                )
                items.append(item)
                by_path[str(path)] = item
                continue
            report = file_service.search(path, request.criteria)
```

(주의: 이 분기는 `try` 블록 안에서 `continue`한다. Python은 `continue` 시에도 `finally`를 실행하므로 루프 끝의 `finally: if on_progress: on_progress(index, total, relative)`가 진행률을 보고한다 — 여기서 `on_progress`를 **직접 호출하지 않는다**(이중 호출 방지).)

`batch_edit`는 **변경하지 않는다**. 이미 `내용 매치 → file_service.apply_edit`(이메일이면 `.md` 산출) → `redact_filename`으로 산출명 정리하는 흐름이 그대로 동작한다. Step 4에서 이를 확인한다.

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv/Scripts/python.exe -m pytest tests/test_batch_service.py -q`
Expected: PASS (신규 3개 + 기존 배치 테스트 유지)

- [ ] **Step 5: Commit**

```bash
git add src/document_redactor/batch_service.py tests/test_batch_service.py
git commit -m "feat: 배치가 .msg/.eml 스캔 + in-place 이메일 내용 스킵·파일명 rename"
```

---

## Task 7: `app.py` UI — 단일 파일 이메일 미리보기/.md 다운로드

**Files:**
- Modify: `app.py`
- Test: `tests/test_app_smoke.py`

**Interfaces:**
- Consumes: `email_service`(직접 호출 없음 — file_service 경유), `FileType.MSG/EML`, `name_redactor`.
- Produces: UI 동작만.

- [ ] **Step 1: 업로더 확장자 + import**

`app.py`의 `st.file_uploader` 호출(단일 파일)의 `type`을 확장:

```python
    uploaded = st.file_uploader(
        "파일 업로드 (.xlsx / .xlsm / 텍스트 PDF / .msg / .eml)",
        type=["xlsx", "xlsm", "pdf", "msg", "eml"],
    )
```

`from document_redactor.models import (...)`에 `FileType`가 이미 포함되어 있는지 확인(포함됨).

- [ ] **Step 2: 단일 파일 이메일 분기 추가**

`render_single_file`에서 `report.total_matches == 0` 분기(이름-only 처리) **다음**, 그리고 기존
“삭제 방식 (Excel만)” 블록(`if report.file_type is FileType.PDF:`) **앞**에 이메일 분기를 삽입한다.
이메일은 삭제 방식 선택 없이 승인 → `.md` 산출한다:

```python
    if report.file_type in (FileType.MSG, FileType.EML):
        st.info("이메일은 서식·이미지·첨부 내용이 보존되지 않는 평문 `.md`로 정리됩니다. "
                "원본 이메일은 수정되지 않습니다.")
        st.dataframe(
            [{"필드": m.field, "줄": m.line, "키워드": m.keyword, "개수": m.count, "문맥": m.context}
             for m in report.email_matches],
            use_container_width=True,
        )
        approved = st.checkbox("위 키워드를 제거한 .md 산출을 승인합니다.", key="s_approve_email")
        if st.button("🗑️ 승인하고 .md 생성", disabled=not approved, type="primary"):
            try:
                request = EditRequest(criteria=report.criteria)
                edit_result = file_service.apply_edit(st.session_state.s_saved, request, OUTPUT_DIR)
                st.session_state.s_edit = edit_result
                st.session_state.s_verify = file_service.verify(Path(edit_result.output_path), report.criteria)
                st.session_state.s_all_selected = True
            except Exception as exc:
                st.error("`.md` 생성 중 오류가 발생했습니다. 결과 파일을 제공하지 않습니다.")
                st.exception(exc)

        edit_result = st.session_state.get("s_edit")
        verification = st.session_state.get("s_verify")
        if edit_result is None or verification is None:
            return

        st.subheader("처리 결과")
        st.metric("제거된 키워드", edit_result.redactions_applied)
        if verification.clean:
            st.success("재검증 완료: 산출 .md에서 키워드가 확인되지 않습니다.")
        else:
            remaining = verification.remaining.total_matches if verification.remaining else 0
            st.error(f"재검증 실패: .md에 키워드가 {remaining}건 남아 있습니다.")

        src_stem = Path(name_redactor.redact_filename(Path(st.session_state.s_saved).name, keywords, case_sensitive)).stem
        dl_name = f"{src_stem}_redacted.md"
        st.download_button(
            "📥 정리된 .md 다운로드",
            data=Path(edit_result.output_path).read_bytes(),
            file_name=dl_name,
            use_container_width=True,
        )
        return
```

- [ ] **Step 3: 앱 import 스모크 테스트**

Run: `.venv/Scripts/python.exe -m pytest tests/test_app_smoke.py -q`
Expected: PASS (문법·import 오류 없음)

- [ ] **Step 4: 전체 테스트**

Run: `.venv/Scripts/python.exe -m pytest -q`
Expected: PASS (전체 그린)

- [ ] **Step 5: Commit**

```bash
git add app.py
git commit -m "feat: 단일 파일 이메일(.msg/.eml) 미리보기 + .md 다운로드"
```

---

## 최종 검증 (수동, 브라우저 미리보기)

- [ ] `.claude/launch.json`의 앱을 띄우고, 키워드에 `포스코` 등을 넣어 `.eml`/`.msg`를 단일 파일로
  업로드 → 미리보기 → `.md` 다운로드를 눈으로 확인한다.
- [ ] 폴더 모드에서 `.eml`이 섞인 폴더를 별도 출력 폴더 모드로 처리해 `<stem>_redacted.md` 산출을 확인한다.

## 미해결/후속 (범위 밖)
- 첨부파일 내용 검사(첨부 엑셀·PDF를 꺼내 기존 서비스로 처리)는 후속.
- HTML 본문 고급 정제(표·링크 보존)는 하지 않음(태그 제거 평문화만).
- `.msg`의 to/cc/첨부까지 포함한 픽스처 생성은 필요 시 후속(현재 순수 EmailContent·.eml로 커버).
