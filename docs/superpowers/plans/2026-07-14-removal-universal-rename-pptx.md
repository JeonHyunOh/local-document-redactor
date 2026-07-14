# 형식 제거·확장자 무관 파일명 정리·pptx 지원 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** (1) 제자리 모드에서 `.dwg/.png/.nwd`를 opt-in 완전 삭제, (2) 제자리 모드 파일명 정리를 확장자 무관 모든 파일로 확장, (3) `.pptx` 내용 키워드 삭제 추가.

**Architecture:** `.pptx`는 excel/pdf와 같은 서비스 계약의 새 `pptx_service`(python-pptx, 문단 단위 검색·재검증). 형식 제거·확장자 무관 rename은 `batch_edit_in_place`에만 추가(우선순위: 제거 > 편집·정리). UI는 서비스만 호출.

**Tech Stack:** Python 3.11+, python-pptx(신규), extract-msg, openpyxl, PyMuPDF, pydantic v2, Streamlit, pytest.

## Global Constraints

- `.dwg/.png/.nwd` 완전 삭제는 **opt-in + 승인 게이트**에서만. 백업 없음, `_removed_log.txt` 기록. 기본 비활성.
- 형식 제거·확장자 무관 rename은 **제자리 모드에만** 적용(출력본 모드 변경 없음).
- pptx 검색·재검증은 **문단(paragraph) 단위**(쪼개진 런 대응). 제거는 런 단위 후 문단 폴백.
- 원본 삭제는 위 예외를 제외하고 금지. 예외를 CLAUDE.md에 명시한다.
- 예외를 `except Exception: pass`로 삼키지 않는다. 파일별 오류는 격리하고 배치는 계속. `pathlib`만.
- 테스트 픽스처 파일은 실행 시 Python으로 생성(레포에 바이너리 금지).

---

## File Structure

- Create: `src/document_redactor/pptx_service.py` — pptx 검색·편집·검증.
- Create: `tests/test_pptx_service.py` — pptx 서비스 테스트.
- Modify: `src/document_redactor/models.py` — `FileType.PPTX`, `PptxMatch`, `SearchReport.pptx_matches`.
- Modify: `src/document_redactor/file_service.py` — `.pptx` 허용·라우팅.
- Modify: `src/document_redactor/batch_service.py` — `.pptx` 스캔, `scan_all_files`, Phase 0 제거, Phase 2a 확장.
- Modify: `tests/conftest.py` — `make_pptx` 픽스처.
- Modify: `tests/test_file_service.py`, `tests/test_batch_service.py` — 라우팅·배치 시나리오.
- Modify: `app.py` — 업로더 pptx, 단일 파일 pptx, 폴더 제거 체크박스.
- Modify: `pyproject.toml` — `python-pptx` 의존성.
- Modify: `CLAUDE.md` — 삭제 예외 명시.

---

## Task 1: 데이터 모델 (PPTX)

**Files:**
- Modify: `src/document_redactor/models.py`
- Test: `tests/test_pptx_service.py`

**Interfaces:**
- Produces: `FileType.PPTX = "pptx"`, `PptxMatch(file_name, slide:int, location:str, keyword:str, count:int, context:str)`, `SearchReport.pptx_matches` + `total_matches` 합산.

- [ ] **Step 1: Write the failing test**

`tests/test_pptx_service.py`(새 파일):

```python
"""pptx_service 및 관련 모델 테스트 — 슬라이드 텍스트 검색·제거·재검증."""

from __future__ import annotations

from pathlib import Path

from document_redactor.models import FileType, PptxMatch, SearchCriteria, SearchReport


def test_filetype_pptx():
    assert FileType.PPTX.value == "pptx"


def test_searchreport_total_includes_pptx():
    report = SearchReport(
        file_name="a.pptx",
        file_type=FileType.PPTX,
        criteria=SearchCriteria(keywords=["x"]),
        pptx_matches=[PptxMatch(file_name="a.pptx", slide=1, location="본문", keyword="x", count=3, context="x x x")],
    )
    assert report.total_matches == 3
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv/Scripts/python.exe -m pytest tests/test_pptx_service.py -q`
Expected: FAIL — `PptxMatch` import 오류.

- [ ] **Step 3: Implement**

`models.py`의 `FileType`에 추가:

```python
    EML = "eml"
    PPTX = "pptx"
    MD = "md"  # 이메일을 정리해 렌더한 Markdown 산출물(입력 유형 아님)
```

`EmailMatch` 아래에 추가:

```python
class PptxMatch(BaseModel):
    """PowerPoint 슬라이드 텍스트 매치 한 건."""

    file_name: str
    slide: int                # 1-기반 슬라이드 번호
    location: str             # "본문"/"표"/"노트"
    keyword: str
    count: int                # 해당 문단에서 발견된 횟수
    context: str              # 해당 문단 텍스트
```

`SearchReport`에 필드·합산 추가:

```python
    email_matches: list[EmailMatch] = Field(default_factory=list)
    pptx_matches: list[PptxMatch] = Field(default_factory=list)
    notes: list[str] = Field(default_factory=list)

    @property
    def total_matches(self) -> int:
        return (
            len(self.excel_matches)
            + sum(m.count for m in self.pdf_matches)
            + sum(m.count for m in self.email_matches)
            + sum(m.count for m in self.pptx_matches)
        )
```

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv/Scripts/python.exe -m pytest tests/test_pptx_service.py -q`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
git add src/document_redactor/models.py tests/test_pptx_service.py
git commit -m "feat: PptxMatch·FileType.PPTX 모델 추가"
```

---

## Task 2: `pptx_service` + 의존성 + 픽스처

**Files:**
- Create: `src/document_redactor/pptx_service.py`
- Modify: `pyproject.toml`, `tests/conftest.py`
- Test: `tests/test_pptx_service.py`

**Interfaces:**
- Consumes: `python-pptx`, `keyword_matcher`, `models`.
- Produces: `search(path, criteria) -> SearchReport`, `apply_edit(path, request, output_dir, selected=None) -> EditResult`, `verify(output_path, criteria) -> VerificationResult`.

- [ ] **Step 1: pyproject + 설치**

`pyproject.toml`의 `dependencies`에 추가:

```toml
    "extract-msg>=0.54",
    "python-pptx>=0.6.23",
```

Run: `.venv/Scripts/python.exe -m pip install -e ".[dev]"`
Expected: `python-pptx` 설치 완료.

- [ ] **Step 2: conftest에 make_pptx 픽스처 추가**

`tests/conftest.py`에 픽스처 추가:

```python
@pytest.fixture
def make_pptx():
    from pptx import Presentation
    from pptx.util import Inches

    def _make(path: Path, *, title="", body_lines=(), table=None, notes="", split_runs=None):
        path.parent.mkdir(parents=True, exist_ok=True)
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(3))
        tf = tb.text_frame
        tf.text = title
        for line in body_lines:
            tf.add_paragraph().text = line
        if split_runs:  # 한 문단을 여러 런으로 분할(쪼개진 키워드 테스트용)
            para = tf.add_paragraph()
            for piece in split_runs:
                run = para.add_run()
                run.text = piece
        if table:
            rows, cols = len(table), len(table[0])
            gt = slide.shapes.add_table(rows, cols, Inches(1), Inches(4.5), Inches(5), Inches(1)).table
            for r, rowvals in enumerate(table):
                for c, val in enumerate(rowvals):
                    gt.cell(r, c).text = val
        if notes:
            slide.notes_slide.notes_text_frame.text = notes
        prs.save(str(path))
        return path

    return _make
```

- [ ] **Step 3: Write the failing test**

`tests/test_pptx_service.py`에 추가:

```python
from document_redactor import pptx_service
from document_redactor.models import EditRequest


def test_search_finds_keywords_across_locations(make_pptx, tmp_path):
    p = make_pptx(
        tmp_path / "d.pptx",
        title="포스코 제목",
        body_lines=["대외비 본문"],
        table=[["포스코 셀", "일반"]],
        notes="노트에 대외비",
    )
    report = pptx_service.search(p, SearchCriteria(keywords=["포스코", "대외비"]))
    assert report.file_type == FileType.PPTX
    assert report.total_matches == 4  # 제목·표(포스코) + 본문·노트(대외비)
    locs = {m.location for m in report.pptx_matches}
    assert {"본문", "표", "노트"} <= locs


def test_apply_edit_removes_keywords_and_keeps_original(make_pptx, tmp_path):
    p = make_pptx(tmp_path / "d.pptx", title="포스코 제목", body_lines=["포스코 대외비"])
    out = tmp_path / "out"
    req = EditRequest(criteria=SearchCriteria(keywords=["포스코"]))
    result = pptx_service.apply_edit(p, req, out)

    produced = Path(result.output_path)
    assert produced.name == "d_edited.pptx"
    assert result.file_type == FileType.PPTX
    # 산출본에 포스코 없음, 원본은 그대로
    assert pptx_service.verify(produced, req.criteria).clean is True
    assert pptx_service.search(p, req.criteria).total_matches > 0


def test_apply_edit_handles_split_runs(make_pptx, tmp_path):
    # '포스'+'코 문서'로 쪼개진 런 → 런 단위 제거는 놓치지만 문단 폴백이 제거
    p = make_pptx(tmp_path / "s.pptx", title="제목", split_runs=["포스", "코 문서"])
    out = tmp_path / "out"
    req = EditRequest(criteria=SearchCriteria(keywords=["포스코"]))
    result = pptx_service.apply_edit(p, req, out)
    assert pptx_service.verify(Path(result.output_path), req.criteria).clean is True


def test_verify_detects_remaining(make_pptx, tmp_path):
    p = make_pptx(tmp_path / "d.pptx", title="포스코 제목")
    verification = pptx_service.verify(p, SearchCriteria(keywords=["포스코"]))
    assert verification.clean is False
    assert verification.remaining is not None
    assert verification.remaining.total_matches == 1
```

- [ ] **Step 4: Run test to verify it fails**

Run: `.venv/Scripts/python.exe -m pytest tests/test_pptx_service.py -q -k "search_finds or apply_edit or verify_detects"`
Expected: FAIL — `pptx_service` 미정의.

- [ ] **Step 5: Implement**

`src/document_redactor/pptx_service.py`:

```python
"""PowerPoint(.pptx) 슬라이드 텍스트의 키워드 검색·제거·재검증.

python-pptx 의존을 이 모듈에 격리한다. 텍스트가 여러 런(run)으로 쪼개질 수 있으므로
검색·재검증은 문단(paragraph) 단위로 합쳐 판정하고, 제거는 런 단위 후 문단에 키워드가
남으면 문단 단위로 폴백해 확실히 제거한다. 원본은 편집하지 않고 <stem>_edited.pptx로 저장한다.
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

from pptx import Presentation

from . import keyword_matcher
from .models import (
    EditRequest,
    EditResult,
    FileType,
    PptxMatch,
    SearchCriteria,
    SearchReport,
    VerificationResult,
)


def _iter_paragraphs(prs) -> Iterator[tuple[int, str, object]]:
    """(슬라이드 1-기반 번호, location, paragraph)를 순회한다. 도형·표·노트를 포함."""
    for slide_index, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    yield slide_index, "본문", para
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            yield slide_index, "표", para
        if slide.has_notes_slide:
            for para in slide.notes_slide.notes_text_frame.paragraphs:
                yield slide_index, "노트", para


def _count(text: str, keywords: list[str], case_sensitive: bool) -> int:
    """문단 텍스트의 키워드(부분문자열) 총 발견 횟수."""
    total = 0
    haystack = text if case_sensitive else text.casefold()
    for keyword in keyword_matcher.normalize_keywords(keywords):
        needle = keyword if case_sensitive else keyword.casefold()
        if needle:
            total += haystack.count(needle)
    return total


def search(path: Path, criteria: SearchCriteria) -> SearchReport:
    """슬라이드·표·노트 문단에서 키워드를 검색한다(파일 무수정)."""
    path = Path(path)
    prs = Presentation(str(path))
    keywords = keyword_matcher.normalize_keywords(criteria.keywords)
    matches: list[PptxMatch] = []
    for slide_index, location, para in _iter_paragraphs(prs):
        text = para.text
        for keyword in keywords:
            n = _count(text, [keyword], criteria.case_sensitive)
            if n:
                matches.append(
                    PptxMatch(
                        file_name=path.name,
                        slide=slide_index,
                        location=location,
                        keyword=keyword,
                        count=n,
                        context=text,
                    )
                )
    return SearchReport(
        file_name=path.name, file_type=FileType.PPTX, criteria=criteria, pptx_matches=matches
    )


def apply_edit(
    path: Path, request: EditRequest, output_dir: Path, selected=None
) -> EditResult:
    """슬라이드 텍스트에서 키워드를 제거해 <stem>_edited.pptx로 저장한다(selected 무시)."""
    path = Path(path)
    prs = Presentation(str(path))
    keywords = request.criteria.keywords
    cs = request.criteria.case_sensitive

    removed = 0
    for _, _, para in _iter_paragraphs(prs):
        before = _count(para.text, keywords, cs)
        if before == 0:
            continue
        removed += before
        # 런 단위 제거
        for run in para.runs:
            run.text = keyword_matcher.remove_keywords(run.text, keywords, cs)
        # 문단에 남으면(쪼개진 런) 첫 런에 정리된 문단 텍스트를 넣고 나머지 런 비우기
        if _count(para.text, keywords, cs) > 0 and para.runs:
            cleaned = keyword_matcher.remove_keywords(para.text, keywords, cs)
            para.runs[0].text = cleaned
            for run in para.runs[1:]:
                run.text = ""

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / f"{path.stem}_edited.pptx"
    prs.save(str(out_path))

    return EditResult(
        source_name=path.name,
        output_path=str(out_path),
        file_type=FileType.PPTX,
        redactions_applied=removed,
        log=[f"슬라이드 텍스트에서 키워드 {removed}건 제거 → {out_path.name}"],
    )


def verify(output_path: Path, criteria: SearchCriteria) -> VerificationResult:
    """저장된 .pptx를 문단 단위로 재검색해 키워드 잔존 여부를 확인한다."""
    remaining = search(Path(output_path), criteria)
    clean = remaining.total_matches == 0
    return VerificationResult(
        output_path=str(output_path), clean=clean, remaining=remaining if not clean else None
    )
```

- [ ] **Step 6: Run test to verify it passes**

Run: `.venv/Scripts/python.exe -m pytest tests/test_pptx_service.py -q`
Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add pyproject.toml tests/conftest.py src/document_redactor/pptx_service.py tests/test_pptx_service.py
git commit -m "feat: pptx_service(문단 단위 검색·런 폴백 제거) + python-pptx 의존성"
```

---

## Task 3: `file_service` 라우팅 (.pptx)

**Files:**
- Modify: `src/document_redactor/file_service.py`
- Test: `tests/test_file_service.py`

**Interfaces:**
- Consumes: `pptx_service`. Produces: `.pptx` 업로드 허용·라우팅.

- [ ] **Step 1: Write the failing test**

`tests/test_file_service.py`에 추가(상단 import에 `pptx_service` 추가):

```python
from document_redactor import email_service, file_service, pptx_service  # 기존 줄 교체


def test_detect_and_route_pptx():
    assert file_service.detect_file_type("a.pptx") is FileType.PPTX
    assert file_service._service_for(Path("a.pptx")) is pptx_service
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv/Scripts/python.exe -m pytest tests/test_file_service.py -q -k "pptx"`
Expected: FAIL — `.pptx` 미지원.

- [ ] **Step 3: Implement**

`file_service.py` 상단 import 교체:

```python
from . import email_service, excel_service, pdf_service, pptx_service
```

`_EXTENSION_MAP`에 추가:

```python
    ".eml": FileType.EML,
    ".pptx": FileType.PPTX,
```

`detect_file_type` 안내 문구 갱신:

```python
            "지원 형식은 .xlsx, .xlsm, .pdf, .msg, .eml, .pptx 입니다. "
```

`_service_for`에 pptx 분기 추가(이메일 분기 뒤):

```python
    if file_type in (FileType.MSG, FileType.EML):
        return email_service
    if file_type is FileType.PPTX:
        return pptx_service
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
git commit -m "feat: file_service가 .pptx 업로드 허용·라우팅"
```

---

## Task 4: `batch_service` — pptx 스캔 + Phase 0 제거 + Phase 2a 확장

**Files:**
- Modify: `src/document_redactor/batch_service.py`
- Test: `tests/test_batch_service.py`

**Interfaces:**
- Consumes: `name_redactor`, `scan_all_files`.
- Produces:
  - `_SUPPORTED_SUFFIXES`에 `.pptx` 포함.
  - `scan_all_files(root, recursive) -> list[Path]`.
  - `batch_edit_in_place(..., remove_suffixes: set[str] | None = None)` — Phase 0 완전 삭제 + Phase 2a 확장자 무관 rename.

- [ ] **Step 1: Write the failing test**

`tests/test_batch_service.py`에 추가:

```python
def test_scan_all_files_returns_every_extension(tmp_path: Path):
    _xlsx(tmp_path / "a.xlsx", "x")
    (tmp_path / "b.dwg").write_bytes(b"dwg")
    (tmp_path / "c.txt").write_text("t", encoding="utf-8")
    (tmp_path / "~$lock.xlsx").write_bytes(b"lock")
    names = {p.name for p in batch_service.scan_all_files(tmp_path, recursive=True)}
    assert names == {"a.xlsx", "b.dwg", "c.txt"}


def test_in_place_removes_target_suffixes_hard_delete(tmp_path: Path):
    root = tmp_path / "src"
    _xlsx(root / "a.xlsx", "대외비 문서")
    (root / "도면.dwg").write_bytes(b"dwg")
    (root / "이미지.png").write_bytes(b"png")
    (root / "모델.nwd").write_bytes(b"nwd")
    backup = tmp_path / "src_backup"
    req = EditRequest(criteria=_criteria("대외비"), excel_action=ExcelAction.CLEAR_CELL)

    items = batch_service.batch_edit_in_place(
        root, req, backup, recursive=True, remove_suffixes={".dwg", ".png", ".nwd"}
    )
    # 완전 삭제됨(백업 없음)
    assert not (root / "도면.dwg").exists()
    assert not (root / "이미지.png").exists()
    assert not (root / "모델.nwd").exists()
    assert not (backup / "도면.dwg").exists()
    # 삭제 로그 기록
    log = (backup / "_removed_log.txt").read_text(encoding="utf-8")
    assert "도면.dwg" in log and "모델.nwd" in log
    # 결과에 note
    assert any(i.note and "완전 삭제" in i.note for i in items)
    # 지원 파일은 정상 편집
    assert load_workbook(root / "a.xlsx").active["A1"].value is None


def test_in_place_removal_off_by_default(tmp_path: Path):
    root = tmp_path / "src"
    _xlsx(root / "a.xlsx", "대외비")
    (root / "도면.dwg").write_bytes(b"dwg")
    backup = tmp_path / "src_backup"
    req = EditRequest(criteria=_criteria("대외비"), excel_action=ExcelAction.CLEAR_CELL)
    batch_service.batch_edit_in_place(root, req, backup, recursive=True)  # remove_suffixes 없음
    assert (root / "도면.dwg").exists()  # 삭제 안 됨


def test_in_place_renames_unsupported_extension_by_filename(tmp_path: Path):
    root = tmp_path / "src"
    (root / "포스코_메모.txt").write_text("공개", encoding="utf-8")  # 미지원 형식, 파일명만 키워드
    backup = tmp_path / "src_backup"
    req = EditRequest(criteria=_criteria("포스코"), excel_action=ExcelAction.CLEAR_CELL)
    items = batch_service.batch_edit_in_place(root, req, backup, recursive=True)

    assert (root / "메모.txt").exists()  # 확장자 무관 파일명 정리
    assert not (root / "포스코_메모.txt").exists()
    assert any(i.renamed_to == "메모.txt" for i in items)


def test_in_place_pptx_content_edited(tmp_path: Path, make_pptx):
    root = tmp_path / "src"
    make_pptx(root / "deck.pptx", title="포스코 제목", body_lines=["대외비 본문"])
    backup = tmp_path / "src_backup"
    req = EditRequest(criteria=_criteria("포스코", "대외비"), excel_action=ExcelAction.CLEAR_CELL)
    batch_service.batch_edit_in_place(root, req, backup, recursive=True)

    from document_redactor import pptx_service
    assert pptx_service.search(root / "deck.pptx", req.criteria).total_matches == 0
    assert (backup / "deck.pptx").exists()  # 원본 백업
```

- [ ] **Step 2: Run test to verify it fails**

Run: `.venv/Scripts/python.exe -m pytest tests/test_batch_service.py -q -k "scan_all or removes_target or removal_off or unsupported_extension or pptx_content"`
Expected: FAIL — `scan_all_files`/`remove_suffixes` 미정의.

- [ ] **Step 3: Implement**

`batch_service.py`의 `_SUPPORTED_SUFFIXES`에 `.pptx` 추가:

```python
_SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".pdf", ".msg", ".eml", ".pptx"}
```

`scan_folder` 함수 아래에 `scan_all_files` 추가:

```python
def scan_all_files(root: Path, recursive: bool = True) -> list[Path]:
    """숨김·Office 잠금(~$) 파일을 제외한 모든 파일(확장자 무관)을 정렬해 반환한다.

    확장자 무관 파일명 정리·형식 제거가 지원 형식에 국한되지 않도록 하는 스캔이다.
    """
    globber = root.rglob("*") if recursive else root.glob("*")
    return sorted(
        p
        for p in globber
        if p.is_file() and not p.name.startswith("~$") and not p.name.startswith(".")
    )
```

`batch_edit_in_place` 시그니처에 `remove_suffixes` 추가:

```python
def batch_edit_in_place(
    root: Path,
    request: EditRequest,
    backup_root: Path,
    recursive: bool = True,
    on_progress: ProgressCallback | None = None,
    remove_suffixes: set[str] | None = None,
) -> list[BatchEditItem]:
```

`keywords`/`cs`/`files`/`total`/`items`/`by_path` 초기화 **직후, `# --- Phase 1` 주석 앞**에
Phase 0을 삽입:

```python
    # --- Phase 0: 형식 제거(opt-in, 완전 삭제·백업 없음) ---
    removed_log: list[str] = []
    if remove_suffixes:
        targets = {s.lower() for s in remove_suffixes}
        for path in scan_all_files(root, recursive):
            if path.suffix.lower() not in targets:
                continue
            rel = path.relative_to(root).as_posix()
            try:
                path.unlink()
                removed_log.append(rel)
                items.append(
                    BatchEditItem(
                        path=str(path),
                        relative_path=rel,
                        note="완전 삭제됨(제거 대상 확장자)",
                    )
                )
            except Exception as exc:  # noqa: BLE001 - 사유 기록 후 계속
                items.append(BatchEditItem(path=str(path), relative_path=rel, error=str(exc)))
    if removed_log:
        backup_root.mkdir(parents=True, exist_ok=True)
        (backup_root / "_removed_log.txt").write_text(
            "\n".join(removed_log) + "\n", encoding="utf-8"
        )
```

Phase 2a를 **확장자 무관**으로 교체. 기존:

```python
    # 2a) 파일 rename — Phase 1에서 실패(error)한 파일은 건드리지 않는다.
    for path in files:
        item = by_path.get(str(path))
        if item is None or item.error:
            continue
        if not name_redactor.name_contains_keyword(path.name, keywords, cs):
            continue
        # 백업 보장(이름-only로 Phase 1에서 백업 안 된 경우)
        backup_path = backup_root / path.relative_to(root)
        if not backup_path.exists():
            backup_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, backup_path)
        new_name = _disk_unique(path.parent, name_redactor.redact_filename(path.name, keywords, cs))
        new_path = path.parent / new_name
        old_rel = path.relative_to(root).as_posix()
        path.rename(new_path)
        renames.append((old_rel, new_path.relative_to(root).as_posix()))
        item.output_path = str(new_path)
        item.renamed_to = new_name
```

를 아래로 교체:

```python
    # 2a) 파일 rename — 확장자 무관 모든 파일. Phase 1 실패(error) 파일은 건드리지 않는다.
    for path in scan_all_files(root, recursive):
        item = by_path.get(str(path))
        if item is not None and item.error:
            continue
        if not name_redactor.name_contains_keyword(path.name, keywords, cs):
            continue
        # 백업 보장(Phase 1에서 백업 안 된 경우)
        backup_path = backup_root / path.relative_to(root)
        if not backup_path.exists():
            backup_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, backup_path)
        new_name = _disk_unique(path.parent, name_redactor.redact_filename(path.name, keywords, cs))
        new_path = path.parent / new_name
        old_rel = path.relative_to(root).as_posix()
        path.rename(new_path)
        renames.append((old_rel, new_path.relative_to(root).as_posix()))
        if item is None:  # 미지원 확장자 파일 — 결과 항목 신규 생성
            item = BatchEditItem(path=str(path), relative_path=old_rel)
            items.append(item)
        by_path[str(new_path)] = item
        item.output_path = str(new_path)
        item.renamed_to = new_name
```

- [ ] **Step 4: Run test to verify it passes**

Run: `.venv/Scripts/python.exe -m pytest tests/test_batch_service.py -q`
Expected: PASS (신규 5개 + 기존 배치 테스트 유지)

- [ ] **Step 5: Commit**

```bash
git add src/document_redactor/batch_service.py tests/test_batch_service.py
git commit -m "feat: 배치 .pptx 스캔 + Phase0 형식 제거 + Phase2a 확장자 무관 rename"
```

---

## Task 5: `app.py` UI — pptx 단일 파일 + 폴더 제거 체크박스

**Files:**
- Modify: `app.py`
- Test: `tests/test_app_smoke.py`

**Interfaces:**
- Consumes: `FileType.PPTX`, `batch_service.batch_edit_in_place(remove_suffixes=...)`.

- [ ] **Step 1: 업로더 + 단일 파일 pptx 분기**

`app.py`의 단일 파일 업로더 `type`에 `"pptx"` 추가:

```python
    uploaded = st.file_uploader(
        "파일 업로드 (.xlsx / .xlsm / 텍스트 PDF / .msg / .eml / .pptx)",
        type=["xlsx", "xlsm", "pdf", "msg", "eml", "pptx"],
    )
```

이메일 분기(`if report.file_type in (FileType.MSG, FileType.EML):`) **바로 뒤**에 pptx 분기 삽입:

```python
    if report.file_type is FileType.PPTX:
        st.info("PowerPoint 슬라이드·표·노트의 텍스트에서 키워드를 제거합니다. 원본은 수정되지 않습니다.")
        st.dataframe(
            [{"슬라이드": m.slide, "위치": m.location, "키워드": m.keyword, "개수": m.count, "문맥": m.context}
             for m in report.pptx_matches],
            use_container_width=True,
        )
        approved = st.checkbox("위 키워드 제거를 승인합니다.", key="s_approve_pptx")
        if st.button("🗑️ 승인하고 수정본 생성", disabled=not approved, type="primary"):
            try:
                request = EditRequest(criteria=report.criteria)
                edit_result = file_service.apply_edit(st.session_state.s_saved, request, OUTPUT_DIR)
                st.session_state.s_edit = edit_result
                st.session_state.s_verify = file_service.verify(Path(edit_result.output_path), report.criteria)
                st.session_state.s_all_selected = True
            except Exception as exc:
                st.error("수정본 생성 중 오류가 발생했습니다. 결과 파일을 제공하지 않습니다.")
                st.exception(exc)

        edit_result = st.session_state.get("s_edit")
        verification = st.session_state.get("s_verify")
        if edit_result is None or verification is None:
            return

        st.subheader("처리 결과")
        st.metric("제거된 키워드", edit_result.redactions_applied)
        if verification.clean:
            st.success("재검증 완료: 산출본에서 키워드가 확인되지 않습니다.")
        else:
            remaining = verification.remaining.total_matches if verification.remaining else 0
            st.error(f"재검증 실패: 키워드가 {remaining}건 남아 있습니다.")

        src_stem = Path(name_redactor.redact_filename(Path(st.session_state.s_saved).name, keywords, case_sensitive)).stem
        st.download_button(
            "📥 수정본 다운로드",
            data=Path(edit_result.output_path).read_bytes(),
            file_name=f"{src_stem}_edited.pptx",
            use_container_width=True,
        )
        return
```

- [ ] **Step 2: 폴더 제자리 모드 — 형식 제거 체크박스**

`render_folder`의 제자리 경고(`if in_place:` 블록) **뒤**에 체크박스를 추가한다. 먼저 상단 상수 정의부(파일 상단, `_EXCEL_ACTION_LABEL` 부근)에 추가:

```python
_REMOVAL_SUFFIXES = {".dwg", ".png", ".nwd"}
```

제자리 경고 블록 다음에:

```python
    remove_targets = False
    if in_place:
        remove_targets = st.checkbox(
            "CAD·이미지·3D 파일(.dwg/.png/.nwd) 완전 삭제 (복구 불가)",
            key="b_remove_targets",
        )
        if remove_targets:
            st.warning("⚠️ 체크한 확장자 파일은 **백업 없이 완전 삭제**됩니다. `_removed_log.txt`에만 목록이 남습니다.")
```

그리고 제자리 실행 호출에 `remove_suffixes`를 전달:

```python
            if in_place:
                st.session_state.b_edit = batch_service.batch_edit_in_place(
                    root, request, backup_root,
                    recursive=st.session_state.b_recursive,
                    on_progress=_on_edit,
                    remove_suffixes=_REMOVAL_SUFFIXES if st.session_state.get("b_remove_targets") else None,
                )
                st.session_state.b_out = None
                st.session_state.b_backup = str(backup_root)
```

(기존 `batch_edit_in_place(...)` 호출을 위 형태로 교체.)

결과 표시부(제자리 성공 안내 부근)에 삭제 로그 안내 추가:

```python
    if in_place_done:
        removed = [e for e in edits if e.note and "완전 삭제" in e.note]
        if removed:
            st.info(f"완전 삭제된 파일: {len(removed)}개 — 목록: `{Path(st.session_state.b_backup) / '_removed_log.txt'}`")
```

(이 블록은 기존 `if in_place_done:` 성공 안내 바로 뒤에 추가한다.)

- [ ] **Step 3: 앱 import 스모크 테스트**

Run: `.venv/Scripts/python.exe -m pytest tests/test_app_smoke.py -q`
Expected: PASS

- [ ] **Step 4: 전체 테스트**

Run: `.venv/Scripts/python.exe -m pytest -q`
Expected: PASS (전체 그린)

- [ ] **Step 5: Commit**

```bash
git add app.py
git commit -m "feat: 단일 파일 pptx UI + 폴더 제자리 형식 제거 체크박스"
```

---

## Task 6: CLAUDE.md 안전 불변식 예외 명시

**Files:**
- Modify: `CLAUDE.md`

- [ ] **Step 1: 삭제 예외 명시**

`CLAUDE.md`의 "폴더 배치 처리" 절에 아래 항목을 추가한다:

```markdown
- `batch_edit_in_place`는 **opt-in + 명시적 승인** 시에만 `.dwg`/`.png`/`.nwd`(내용 정리 불가 형식)를
  **완전 삭제**할 수 있다(백업 없음, `backup_root/_removed_log.txt`에 목록 기록, 복구 불가).
  기본값은 비활성이며, 이 경로 외에는 "원본 파일을 절대 삭제하지 않는다" 불변식을 그대로 지킨다.
- 제자리 모드의 파일명 키워드 정리는 **확장자 무관 모든 파일**에 적용한다(내용 편집은 지원 형식만).
```

- [ ] **Step 2: Commit**

```bash
git add CLAUDE.md
git commit -m "docs: 형식 완전 삭제(opt-in)·확장자 무관 파일명 정리 불변식 명시"
```

---

## 최종 검증 (수동, 브라우저 미리보기)

- [ ] 앱을 띄우고 `.pptx`를 단일 업로드 → 미리보기 → `_edited.pptx` 다운로드 확인.
- [ ] `.dwg`/`.png`/`.nwd`와 지원 파일이 섞인 **폴더 사본**으로 제자리 모드 + 삭제 체크박스 → 완전
  삭제·`_removed_log.txt`·확장자 무관 파일명 정리를 확인한다(반드시 사본으로 시험).

## 미해결/후속 (범위 밖)
- 출력본 모드의 형식 제거·확장자 무관 rename.
- pptx 그룹 도형 깊은 재귀·차트/SmartArt 텍스트.
