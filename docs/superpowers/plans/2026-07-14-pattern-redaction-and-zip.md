# 패턴 자동 삭제 + ZIP 처리 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:executing-plans. Steps use checkbox syntax.

**Goal:** (A) 키워드와 무관하게 이메일·주민번호·카드·전화번호를 항상 제거(계좌번호는 opt-in), (B) 제자리 모드에서 zip을 해제해 내부를 정리하고 원본 zip 삭제.

**Architecture:** 순수 모듈 `pattern_matcher`를 모든 콘텐츠 서비스(excel/pdf/pptx/email)의 search·edit·verify에 통합. zip은 `batch_edit_in_place`의 새 Phase Z(가장 먼저)로 해제.

**Tech Stack:** Python 3.11+, 표준 `re`·`zipfile`, 기존 스택.

## Global Constraints
- 항상 적용 패턴: 이메일·주민등록번호·신용카드·전화번호. 계좌번호는 `SearchCriteria.redact_account_numbers`(기본 False)로만.
- 패턴 제거는 항상 부분문자열 제거. 재검증에 패턴 포함.
- zip은 제자리 모드 전용, 원본 zip은 백업 후 삭제, zip-slip 방지, 중첩 반복 해제.
- 원본 삭제 예외(zip)도 backup 보장. 예외 삼키지 않고 파일별 격리. pathlib만.
- 테스트 픽스처는 실행 시 생성.

---

## Task A1: `pattern_matcher` 순수 모듈

**Files:** Create `src/document_redactor/pattern_matcher.py`, `tests/test_pattern_matcher.py`

**Interfaces:**
- `find_patterns(text, include_account=False) -> list[tuple[str,str]]`  # (유형, 매칭값)
- `remove_patterns(text, include_account=False) -> str`
- `contains_pattern(text, include_account=False) -> bool`

- [ ] Step1: Write failing tests

```python
"""pattern_matcher 테스트 — 정형 개인정보 패턴 감지·제거·오탐 경계."""
from document_redactor import pattern_matcher as pm

def test_finds_always_on_types():
    text = "연락 010-1234-5678, 메일 a@b.com, 주민 900101-1234567, 카드 1234-5678-9012-3456"
    types = {t for t, _ in pm.find_patterns(text)}
    assert types == {"전화번호", "이메일", "주민등록번호", "신용카드"}

def test_account_is_opt_in():
    text = "계좌 123-45-678901"
    assert pm.find_patterns(text) == []               # 기본은 계좌 제외
    assert any(t == "계좌번호" for t, _ in pm.find_patterns(text, include_account=True))

def test_no_false_positive_on_plain_numbers():
    assert pm.find_patterns("제품 A-100, 일반숫자 123456") == []

def test_remove_patterns_strips_matches():
    out = pm.remove_patterns("메일 a@b.com 끝")
    assert "a@b.com" not in out and "메일" in out and "끝" in out

def test_remove_overlapping_prefers_longer():
    # 카드번호가 전화번호보다 길게 매칭되어 전체 제거
    out = pm.remove_patterns("카드 1234-5678-9012-3456 끝")
    assert not any(ch.isdigit() for ch in out)

def test_contains_pattern():
    assert pm.contains_pattern("a@b.com")
    assert not pm.contains_pattern("그냥 텍스트")
```

- [ ] Step2: Run — FAIL (module 없음)
- [ ] Step3: Implement `pattern_matcher.py`

```python
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
    "계좌번호", re.compile(r"\b\d{2,6}-\d{2,6}-\d{2,6}(?:-\d{1,6})?\b")
)


def _active(include_account: bool) -> list[tuple[str, re.Pattern]]:
    return _ALWAYS + ([_ACCOUNT] if include_account else [])


def _spans(text: str, include_account: bool) -> list[tuple[int, int, str, str]]:
    """(start, end, 유형, 매칭값) 목록. 겹치면 긴 매치 우선으로 비겹침 선택."""
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
    return bool(_spans(text, include_account))
```

- [ ] Step4: Run — PASS
- [ ] Step5: Commit `feat: pattern_matcher 순수 모듈(정형 개인정보 패턴 감지·제거)`

---

## Task A2: 모델 — `SearchCriteria.redact_account_numbers`

**Files:** Modify `src/document_redactor/models.py`, test in `tests/test_pattern_matcher.py` 또는 model 테스트.

- [ ] Step1: Add field (기존 필드 유지)

```python
class SearchCriteria(BaseModel):
    keywords: list[str] = Field(default_factory=list)
    mode: SearchMode = SearchMode.CONTAINS
    case_sensitive: bool = False
    redact_account_numbers: bool = False  # 계좌번호 패턴 제거 opt-in(오탐 위험)
```

- [ ] Step2: 기존 테스트 통과 확인 → Commit `feat: SearchCriteria.redact_account_numbers 추가`

---

## Task A3: excel_service 패턴 통합

**Files:** Modify `src/document_redactor/excel_service.py`, `tests/test_excel_service.py`

**Interfaces:** search/edit/verify가 키워드와 함께 패턴도 처리.

- [ ] Step1: Write failing tests (키워드 없이 패턴만으로도 검출·제거)

```python
def test_excel_search_detects_pattern_without_keyword(tmp_path):
    from openpyxl import Workbook
    from document_redactor import excel_service
    from document_redactor.models import SearchCriteria
    p = tmp_path/"a.xlsx"; wb=Workbook(); wb.active["A1"]="연락 010-1234-5678"; wb.save(p)
    report = excel_service.search(p, SearchCriteria(keywords=[]))  # 키워드 없음
    assert report.total_matches >= 1
    assert any("전화번호" in m.keyword for m in report.excel_matches)

def test_excel_edit_removes_pattern(tmp_path):
    from openpyxl import Workbook, load_workbook
    from document_redactor import excel_service
    from document_redactor.models import SearchCriteria, EditRequest, ExcelAction
    p = tmp_path/"a.xlsx"; wb=Workbook(); wb.active["A1"]="메일 a@b.com"; wb.save(p)
    req = EditRequest(criteria=SearchCriteria(keywords=[]), excel_action=ExcelAction.REMOVE_KEYWORD)
    res = excel_service.apply_edit(p, req, tmp_path/"out")
    assert excel_service.verify(__import__("pathlib").Path(res.output_path), req.criteria).clean
```

- [ ] Step2: Run — FAIL
- [ ] Step3: Implement — `_searchable_text` 대상 셀에 대해 키워드 매치 + `pattern_matcher.find_patterns` 병합.

핵심 변경(개념):
- `search`: 각 셀 text에서 `find_matches`(키워드) + `find_patterns`(패턴, `criteria.redact_account_numbers`)를 합쳐 ExcelMatch 생성. 패턴은 `keyword=f"[{유형}]"`, `original_value=text`.
- `_apply_all`/`_apply_selected`(REMOVE_KEYWORD): `remove_keywords` 후 `remove_patterns`도 적용. CLEAR_CELL/DELETE_ROW은 "키워드 OR 패턴" 매치 셀을 대상으로.
- 매치 판정 헬퍼 추가:
```python
from . import keyword_matcher, pattern_matcher
def _cell_hits(text, criteria):
    kws = keyword_matcher.find_matches(text, criteria)
    pats = pattern_matcher.find_patterns(text, criteria.redact_account_numbers)
    return kws, pats
def _redact_text(text, criteria):
    t = keyword_matcher.remove_keywords(text, criteria.keywords, criteria.case_sensitive)
    return pattern_matcher.remove_patterns(t, criteria.redact_account_numbers)
```
`search`/`_apply_*`에서 `find_matches`→`_cell_hits`, `remove_keywords`→`_redact_text`로 대체하고, 판정 조건을 `if kws or pats`로 확장.

- [ ] Step4: Run — PASS (기존 엑셀 테스트도 유지: 패턴 없는 텍스트는 영향 없음)
- [ ] Step5: Commit `feat: excel_service에 패턴 자동 삭제 통합`

---

## Task A4: pptx_service · email_service 패턴 통합

**Files:** Modify `pptx_service.py`, `email_service.py`, 각 테스트.

- [ ] Step1: Failing tests (키워드 없이 패턴 검출·제거·verify clean) — 각 서비스에 1~2개.
- [ ] Step2: Run — FAIL
- [ ] Step3: Implement
  - 공통 `_count`/검색/제거 지점에 패턴을 병합. 두 서비스 모두 텍스트 기반이라:
    - 검색: 문단/렌더 텍스트에 `pattern_matcher.find_patterns(text, criteria.redact_account_numbers)` 추가 → 매치 생성(`keyword=f"[{유형}]"`).
    - 제거: `remove_keywords` 뒤 `remove_patterns` 적용. (pptx는 런 단위 후 문단 폴백도 패턴 포함해서 판정.)
    - verify: `search`가 패턴 포함하므로 자동 반영.
  - pptx `_count`류 잔존 판정은 "키워드 수 + 패턴 매치 수"로 계산하도록 헬퍼 확장.
- [ ] Step4: Run — PASS
- [ ] Step5: Commit `feat: pptx·email 서비스에 패턴 자동 삭제 통합`

---

## Task A5: pdf_service 패턴 통합

**Files:** Modify `pdf_service.py`, `tests/test_pdf_service.py`

- [ ] Step1: Failing test — 텍스트 PDF에 이메일/전화 넣고(ASCII), 키워드 없이 redaction 후 잔존 0.
- [ ] Step2: Run — FAIL
- [ ] Step3: Implement
  - `search`: 페이지 `get_text()`에서 `pattern_matcher.find_patterns`로 매칭값을 얻고, 각 값을 `_page_search`(search_for)로 위치 잡아 PdfMatch 생성(`keyword=f"[{유형}]"`, count=발견 rect 수).
  - `apply_edit`: 키워드 redaction과 동일하게, 패턴 매칭값도 `search_for` → `add_redact_annot`. selected 미지정(배치·always) 시 키워드 + 패턴 모두.
  - 레이아웃 공백으로 search_for가 매칭값을 못 찾을 수 있음 → note 유지.
- [ ] Step4: Run — PASS
- [ ] Step5: Commit `feat: pdf_service에 패턴 자동 삭제(get_text→search_for redaction) 통합`

---

## Task A6: app.py — 계좌 opt-in 체크박스 + 안내

**Files:** Modify `app.py`, `tests/test_app_smoke.py`

- [ ] Step1: 사이드바에 안내 + 체크박스 추가

```python
    st.divider()
    st.caption("개인정보 패턴(이메일·주민번호·카드·전화)은 키워드와 무관하게 항상 자동 삭제됩니다.")
    st.checkbox("계좌번호 패턴도 삭제 (⚠️ 일반 숫자열 오탐 가능)", key="redact_account")
    redact_account = st.session_state.get("redact_account", False)
```

`build_criteria()`가 `redact_account_numbers=redact_account`를 넘기도록 수정.

- [ ] Step2: 스모크 통과 → Commit `feat: 계좌번호 패턴 opt-in 체크박스 + 안내`

---

## Task B1: ZIP 해제 Phase Z (제자리 모드)

**Files:** Modify `src/document_redactor/batch_service.py`, `tests/test_batch_service.py`

**Interfaces:** `batch_edit_in_place`가 시작 시 zip을 해제(백업 후 삭제)하고, 이후 단계가 해제된 파일을 처리.

- [ ] Step1: Failing tests

```python
import zipfile

def _make_zip(path, inner_name, data: bytes):
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr(inner_name, data)

def test_in_place_extracts_zip_and_deletes_original(tmp_path, make_eml):
    root = tmp_path/"src"; root.mkdir(parents=True)
    from openpyxl import Workbook
    import io
    buf = io.BytesIO(); wb=Workbook(); wb.active["A1"]="대외비"; wb.save(buf)
    _make_zip(root/"압축.zip", "포스코_내부.xlsx", buf.getvalue())
    backup = tmp_path/"src_backup"
    req = EditRequest(criteria=_criteria("대외비"), excel_action=ExcelAction.CLEAR_CELL)
    batch_service.batch_edit_in_place(root, req, backup, recursive=True)

    assert not (root/"압축.zip").exists()               # 원본 zip 삭제
    assert (backup/"압축.zip").exists()                  # 백업 보존
    # 해제된 폴더 안 파일이 정리됨(내용 편집 + 파일명 정리)
    extracted = root/"압축"/"내부.xlsx"
    assert extracted.exists()
    from openpyxl import load_workbook
    assert load_workbook(extracted).active["A1"].value is None

def test_in_place_zip_slip_rejected(tmp_path):
    root = tmp_path/"src"; root.mkdir(parents=True)
    _make_zip(root/"evil.zip", "../escape.txt", b"x")
    backup = tmp_path/"src_backup"
    req = EditRequest(criteria=_criteria("x"), excel_action=ExcelAction.CLEAR_CELL)
    items = batch_service.batch_edit_in_place(root, req, backup, recursive=True)
    assert not (tmp_path/"escape.txt").exists()          # 상위로 탈출 안 됨
    assert any(i.error for i in items if i.path.endswith("evil.zip"))
```

- [ ] Step2: Run — FAIL
- [ ] Step3: Implement — 상수·헬퍼·Phase Z 추가

```python
import zipfile

def _safe_extract_zip(zip_path: Path, dest_dir: Path) -> None:
    """zip을 dest_dir로 안전하게 해제한다(zip-slip 방지). 실패 시 예외."""
    with zipfile.ZipFile(zip_path) as zf:
        dest_dir.mkdir(parents=True, exist_ok=True)
        base = dest_dir.resolve()
        for member in zf.namelist():
            target = (dest_dir / member).resolve()
            if base != target and base not in target.parents:
                raise ValueError(f"안전하지 않은 zip 경로(zip-slip): {member}")
        zf.extractall(dest_dir)
```

`batch_edit_in_place` 맨 앞(Phase 0 이전)에 Phase Z 추가. `.zip`이 없어질 때까지 반복:

```python
    # --- Phase Z: zip 해제(제자리 전용) — 원본 zip 백업 후 삭제, 중첩 반복 ---
    while True:
        zips = [p for p in scan_all_files(root, recursive) if p.suffix.lower() == ".zip"]
        if not zips:
            break
        progressed = False
        for zpath in zips:
            rel = zpath.relative_to(root).as_posix()
            try:
                dest = zpath.parent / zpath.stem
                if dest.exists():
                    dest = zpath.parent / _disk_unique(zpath.parent, zpath.stem)
                _safe_extract_zip(zpath, dest)
                backup_path = backup_root / zpath.relative_to(root)
                backup_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(zpath, backup_path)
                zpath.unlink()
                items.append(BatchEditItem(path=str(zpath), relative_path=rel,
                    note="압축 해제 후 원본 zip 삭제"))
                progressed = True
            except Exception as exc:  # noqa: BLE001
                items.append(BatchEditItem(path=str(zpath), relative_path=rel, error=str(exc)))
        if not progressed:
            break  # 모든 zip이 실패(손상 등) → 무한루프 방지
```

주의: `_disk_unique`는 파일명 충돌용이지만 폴더명에도 사용 가능(동일 로직). Phase Z가 `items`/`by_path` 초기화 이후에 오도록 배치(기존 `items: list = []` 다음).

`.zip`을 `scan_folder`(내용 처리 대상)에는 넣지 않는다 — zip 자체는 내용 편집 대상이 아니라 해제 대상.

- [ ] Step4: Run — PASS (전체 배치 테스트 유지)
- [ ] Step5: Commit `feat: 제자리 모드 zip 해제 Phase(백업 후 삭제·zip-slip 방지·중첩 반복)`

---

## Task B2: 문서 갱신 + 전체 검증

- [ ] `CLAUDE.md`에 패턴 자동 삭제·zip 처리 규칙 추가.
- [ ] `.venv/Scripts/python.exe -m pytest -q` 전체 그린.
- [ ] Commit `docs: 패턴 자동 삭제·zip 처리 규칙 명시`

## 최종 검증(수동)
- [ ] 앱에서 키워드 없이 전화/이메일 든 파일 → 자동 제거 확인.
- [ ] 폴더 사본으로 제자리 모드 + zip 포함 → 해제·정리·zip 삭제 확인.

## 미해결/후속
- 출력본 모드 zip, 암호 zip, 카드 Luhn 검증.
