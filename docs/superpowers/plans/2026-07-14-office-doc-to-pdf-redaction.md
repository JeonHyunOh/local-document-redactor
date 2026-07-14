# 오피스 문서 → PDF 변환 후 검사·삭제 Implementation Plan

> REQUIRED SUB-SKILL: superpowers:executing-plans. 체크박스 단계.

**Goal:** docx/doc/hwp/hwpx를 PDF로 변환(Word/한글 COM)한 뒤 pdf_service로 키워드·패턴 redaction. 결과 `<stem>_redacted.pdf`, 제자리 모드는 원본 백업 후 삭제.

**Architecture:** COM은 `doc_converter`에 격리, `office_service`가 변환→pdf_service 위임. `file_service`/`batch_service`/`app.py`가 라우팅·배치·UI.

## Global Constraints
- pywin32 + Windows + Word/한글 필수. 미충족·변환실패·보호문서 → ConversionError로 격리.
- 결과는 PDF. 제자리 모드만 원본 삭제(백업 후). 변환 산출·임시파일 정리.
- doc_converter는 monkeypatch 가능 → office_service는 실제 Office 없이 단위 테스트.

---

## Task 1: 모델 — FileType 4종
- Modify `models.py`: `FileType`에 `DOCX="docx"`, `DOC="doc"`, `HWP="hwp"`, `HWPX="hwpx"`.
- [ ] Test(신규 test_office_service.py 상단): `FileType.DOCX.value=="docx"` 등. Run→FAIL→구현→PASS→commit.

---

## Task 2: doc_converter (COM) + pywin32
- Create `doc_converter.py`, `tests/test_doc_converter.py`. Modify `pyproject.toml`(`pywin32; sys_platform=='win32'`).

- [ ] Step1: pyproject에 의존성 추가 + 설치
```toml
    "python-pptx>=0.6.23",
    "pywin32>=305; sys_platform == 'win32'",
```
Run: `.venv/Scripts/python.exe -m pip install -e ".[dev]"`

- [ ] Step2: Write `doc_converter.py`
```python
"""오피스 문서(docx/doc/hwp/hwpx)를 PDF로 변환(Word/한글 COM 자동화, pywin32).

COM 의존을 이 모듈에만 둔다. 변환 결과 PDF는 임시 폴더에 캐시하며, 같은 원본(경로·크기·mtime)
은 재변환하지 않는다(검사→편집 이중 변환 방지). 변환 뒤 앱을 종료하고 실패는 ConversionError로
격리한다. Windows + 해당 오피스 설치가 필요하다.
"""
from __future__ import annotations
import tempfile
from pathlib import Path

_WORD_SUFFIXES = {".docx", ".doc"}
_HWP_SUFFIXES = {".hwp", ".hwpx"}
SUPPORTED_SUFFIXES = _WORD_SUFFIXES | _HWP_SUFFIXES

_cache: dict[tuple[str, int, float], Path] = {}
_cache_dir: Path | None = None


class ConversionError(RuntimeError):
    """오피스 문서를 PDF로 변환하지 못했을 때 발생(환경·보호문서·엔진 오류)."""


def _out_dir() -> Path:
    global _cache_dir
    if _cache_dir is None:
        _cache_dir = Path(tempfile.mkdtemp(prefix="doc2pdf_"))
    return _cache_dir


def convert_to_pdf(path: Path) -> Path:
    """오피스 문서를 PDF로 변환해 그 경로를 반환한다(캐시 사용)."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise ConversionError(f"변환 대상이 아닙니다: {suffix}")
    stat = path.stat()
    key = (str(path.resolve()), stat.st_size, stat.st_mtime)
    cached = _cache.get(key)
    if cached and cached.exists():
        return cached
    out_pdf = _out_dir() / f"{path.stem}.pdf"
    if suffix in _WORD_SUFFIXES:
        _convert_word(path, out_pdf)
    else:
        _convert_hwp(path, out_pdf)
    if not out_pdf.exists():
        raise ConversionError(f"변환 결과 PDF가 생성되지 않았습니다: {path.name}")
    _cache[key] = out_pdf
    return out_pdf


def _convert_word(src: Path, out_pdf: Path) -> None:
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:  # noqa: BLE001
        raise ConversionError("pywin32/COM을 사용할 수 없습니다(Windows+Word 필요).") from exc
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(src), ReadOnly=True)
        doc.SaveAs(str(out_pdf), 17)  # wdFormatPDF
        doc.Close(False)
    except Exception as exc:  # noqa: BLE001
        raise ConversionError(f"Word 변환 실패: {src.name} — {exc}") from exc
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:  # noqa: BLE001
                pass
        pythoncom.CoUninitialize()


def _convert_hwp(src: Path, out_pdf: Path) -> None:
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:  # noqa: BLE001
        raise ConversionError("pywin32/COM을 사용할 수 없습니다(Windows+한글 필요).") from exc
    pythoncom.CoInitialize()
    hwp = None
    try:
        hwp = win32.Dispatch("HwpFrame.HwpObject")
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:  # noqa: BLE001 - 보안모듈 없으면 무시(대화상자 가능)
            pass
        hwp.Open(str(src))
        hwp.SaveAs(str(out_pdf), "PDF", "")
    except Exception as exc:  # noqa: BLE001
        raise ConversionError(f"한글 변환 실패: {src.name} — {exc}") from exc
    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:  # noqa: BLE001
                pass
        pythoncom.CoUninitialize()
```

- [ ] Step3: `tests/test_doc_converter.py` — 실제 변환은 가용 시에만
```python
import pytest
from pathlib import Path
from document_redactor import doc_converter


def _word_available() -> bool:
    try:
        import pythoncom, win32com.client as win32
        pythoncom.CoInitialize()
        try:
            w = win32.Dispatch("Word.Application"); w.Quit(); return True
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        return False


def test_unsupported_suffix_raises(tmp_path):
    p = tmp_path / "a.txt"; p.write_text("x")
    with pytest.raises(doc_converter.ConversionError):
        doc_converter.convert_to_pdf(p)


@pytest.mark.skipif(not _word_available(), reason="Word COM 미가용")
def test_word_docx_to_pdf(tmp_path):
    import win32com.client as win32, pythoncom
    docx = tmp_path / "t.docx"
    pythoncom.CoInitialize()
    w = win32.Dispatch("Word.Application"); w.Visible = False
    d = w.Documents.Add(); d.Content.Text = "대외비 문서"; d.SaveAs(str(docx), 16); d.Close(False); w.Quit()
    pythoncom.CoUninitialize()
    pdf = doc_converter.convert_to_pdf(docx)
    import fitz
    assert "대외비" in fitz.open(str(pdf))[0].get_text()
```
Run: `pytest tests/test_doc_converter.py -q` → PASS(또는 skip). Commit.

---

## Task 3: office_service (search/apply_edit/verify)
- Create `office_service.py`, `tests/test_office_service.py`.
- office_service는 `doc_converter.convert_to_pdf`를 호출하고 `pdf_service`에 위임.

- [ ] Step1: 실패 테스트(변환기 monkeypatch → 가짜 PDF)
```python
from pathlib import Path
import fitz
from document_redactor import office_service, doc_converter
from document_redactor.models import FileType, SearchCriteria, EditRequest


def _fake_pdf(tmp, text):
    p = tmp / "conv.pdf"; d = fitz.open(); pg = d.new_page(); pg.insert_text((72,72), text); d.save(str(p)); d.close(); return p


def test_office_search_and_edit(tmp_path, monkeypatch):
    src = tmp_path / "보고서.docx"; src.write_bytes(b"stub")
    monkeypatch.setattr(doc_converter, "convert_to_pdf", lambda p: _fake_pdf(tmp_path, "대외비 010-1234-5678"))
    rep = office_service.search(src, SearchCriteria(keywords=["대외비"]))
    assert rep.total_matches >= 2  # 키워드 + 전화 패턴
    out = tmp_path / "out"
    res = office_service.apply_edit(src, EditRequest(criteria=SearchCriteria(keywords=["대외비"])), out)
    assert Path(res.output_path).name == "보고서_redacted.pdf"
    assert office_service.verify(Path(res.output_path), SearchCriteria(keywords=["대외비"])).clean
```

- [ ] Step2: Run→FAIL. Step3: Implement
```python
"""오피스 문서를 PDF로 변환(doc_converter)한 뒤 pdf_service로 검사·redaction하는 서비스.

원본은 편집하지 않고, 정리 결과는 <원본stem>_redacted.pdf로 산출한다. 변환은 doc_converter에
격리되어 있어 이 서비스 로직은 변환기를 대체(monkeypatch)해 테스트할 수 있다.
"""
from __future__ import annotations
from pathlib import Path
from . import doc_converter, pdf_service
from .models import EditRequest, EditResult, SearchCriteria, SearchReport, VerificationResult

_RENDER_NOTE = "원본 문서를 PDF로 변환해 검사·정리했습니다(결과는 PDF)."


def search(path: Path, criteria: SearchCriteria) -> SearchReport:
    path = Path(path)
    pdf = doc_converter.convert_to_pdf(path)
    report = pdf_service.search(pdf, criteria)
    report.file_name = path.name
    if _RENDER_NOTE not in report.notes:
        report.notes.append(_RENDER_NOTE)
    return report


def apply_edit(path: Path, request: EditRequest, output_dir: Path, selected=None) -> EditResult:
    path = Path(path)
    pdf = doc_converter.convert_to_pdf(path)
    edit = pdf_service.apply_edit(pdf, request, output_dir)  # <conv>_redacted.pdf
    produced = Path(edit.output_path)
    final = produced.with_name(f"{path.stem}_redacted.pdf")
    if final != produced:
        if final.exists():
            final.unlink()
        produced.rename(final)
    edit.output_path = str(final)
    edit.source_name = path.name
    return edit


def verify(output_path: Path, criteria: SearchCriteria) -> VerificationResult:
    return pdf_service.verify(Path(output_path), criteria)
```
- [ ] Step4: Run→PASS. Commit.

---

## Task 4: file_service 라우팅
- Modify `file_service.py`, `tests/test_file_service.py`.
- [ ] `_EXTENSION_MAP`에 `.docx/.doc/.hwp/.hwpx` → FileType.*. `detect_file_type` 문구 갱신.
- [ ] `_service_for`: 네 형식 → `office_service`(import 추가).
- [ ] Test: `detect_file_type("a.docx") is FileType.DOCX`, `_service_for(Path("a.hwp")) is office_service`. Run→FAIL→구현→PASS. Commit.

---

## Task 5: batch_service — 스캔 + 제자리 오피스 처리
- Modify `batch_service.py`, `tests/test_batch_service.py`.
- [ ] `_SUPPORTED_SUFFIXES`에 네 형식 추가. 상수 `_OFFICE_SUFFIXES = {".docx",".doc",".hwp",".hwpx"}`.
- [ ] `batch_edit_in_place` Phase 1에 **이메일 분기와 같은 패턴의 오피스 분기** 추가(이메일 분기 바로 뒤):
```python
            if path.suffix.lower() in _OFFICE_SUFFIXES:
                report = file_service.search(path, request.criteria)
                name_hit = name_redactor.name_contains_keyword(path.name, keywords, cs)
                if report.total_matches == 0 and not name_hit:
                    item = BatchEditItem(path=str(path), relative_path=relative)
                    items.append(item); by_path[str(path)] = item; continue
                with tempfile.TemporaryDirectory(prefix="redactor_office_") as tmp:
                    edit = file_service.apply_edit(path, request, Path(tmp))
                    pdf_tmp = Path(edit.output_path)
                    verification = file_service.verify(pdf_tmp, request.criteria)
                    if not verification.clean:
                        item = BatchEditItem(path=str(path), relative_path=relative, edit=edit,
                            verification=verification, error="변환·정리본 재검증 실패로 원본 유지(.pdf 미생성).")
                        items.append(item); by_path[str(path)] = item; continue
                    backup_path = backup_root / path.relative_to(root)
                    backup_path.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(path, backup_path)
                    clean_stem = Path(name_redactor.redact_filename(path.name, keywords, cs)).stem
                    pdf_name = _disk_unique(path.parent, f"{clean_stem}_redacted.pdf")
                    pdf_final = path.parent / pdf_name
                    shutil.move(str(pdf_tmp), str(pdf_final))
                path.unlink()
                item = BatchEditItem(path=str(path), relative_path=relative, output_path=str(pdf_final),
                    edit=edit, verification=verification,
                    renamed_to=pdf_final.relative_to(root).as_posix(),
                    note="문서를 PDF로 변환·정리, 원본 삭제")
                items.append(item); by_path[str(path)] = item; continue
```
- [ ] Test(변환기 monkeypatch): 제자리 모드에서 docx→`<stem>_redacted.pdf` 생성 + 원본 삭제 + 백업. `batch_edit`(출력본)은 별도 코드 없이 동작(스캔만 추가) — 테스트로 확인.
  - monkeypatch 대상: `document_redactor.doc_converter.convert_to_pdf` → 가짜 PDF 반환.
- Run→PASS. Commit.

---

## Task 6: app.py UI
- Modify `app.py`.
- [ ] 업로더 `type`에 `docx,doc,hwp,hwpx` 추가. 안내: "문서는 PDF로 변환 후 검사되며 변환에 시간이 걸릴 수 있습니다."
- [ ] 단일 파일: 오피스 형식이면 변환→검사(pdf 매치 표)→승인→`<stem>_redacted.pdf` 다운로드. (pdf 분기와 동일 UI를 재사용하거나 전용 분기.)
  - 간단화: office 형식은 `report.file_type`이 PDF(변환 결과)로 오므로, **기존 PDF/Excel 분기 앞에 office 전용 분기**를 두어 변환 안내 + apply_edit + `.pdf` 다운로드.
- [ ] 스모크 통과. Commit.

---

## Task 7: 문서 + 전체 검증
- [ ] `CLAUDE.md`에 오피스 변환 규칙(Windows+Office 필수, 결과 PDF, 제자리 원본 삭제) 명시.
- [ ] `pytest -q` 전체 그린(변환 실테스트는 이 PC에서만, 그 외 skip).
- [ ] 최종 수동: 실제 .docx/.hwp로 단일·폴더 확인.
- Commit.

## 미해결/후속
- 변환 인스턴스 재사용(배치 성능), 비밀번호 문서, LibreOffice 대체.
