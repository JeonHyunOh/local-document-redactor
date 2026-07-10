"""pdf_service 단위 테스트 (M3). 픽스처는 런타임에 PyMuPDF로 생성한다."""

from __future__ import annotations

from pathlib import Path

import fitz
import pytest

from document_redactor import pdf_service
from document_redactor.models import EditRequest, PdfAction, SearchCriteria, SearchMode


def _make_pdf(path: Path, pages: list[str]) -> Path:
    """각 문자열을 한 페이지로 담은 텍스트 PDF를 생성한다."""
    doc = fitz.open()
    for text in pages:
        page = doc.new_page()
        page.insert_text((72, 72), text, fontsize=14)
    doc.save(path)
    doc.close()
    return path


def _criteria(*keywords: str, mode: SearchMode = SearchMode.CONTAINS) -> SearchCriteria:
    return SearchCriteria(keywords=list(keywords), mode=mode, case_sensitive=False)


@pytest.fixture
def sample_pdf(tmp_path: Path) -> Path:
    return _make_pdf(
        tmp_path / "sample.pdf",
        [
            "This is CONFIDENTIAL page one CONFIDENTIAL",  # 같은 페이지 2회
            "Second page mentions Secret once",
            "Third page has nothing relevant",
        ],
    )


# --------------------------------------------------------------------------- #
# 검색
# --------------------------------------------------------------------------- #
def test_search_multiple_pages(sample_pdf: Path):
    report = pdf_service.search(sample_pdf, _criteria("CONFIDENTIAL", "Secret"))
    pages = {m.page for m in report.pdf_matches}
    assert pages == {1, 2}


def test_same_page_multiple_hits_counts(sample_pdf: Path):
    report = pdf_service.search(sample_pdf, _criteria("CONFIDENTIAL"))
    hit = next(m for m in report.pdf_matches if m.page == 1)
    assert hit.count == 2
    assert hit.rects and len(hit.rects) == 2


def test_search_provides_context(sample_pdf: Path):
    report = pdf_service.search(sample_pdf, _criteria("Secret"))
    hit = next(m for m in report.pdf_matches if m.keyword == "Secret")
    assert hit.context and "Secret" in hit.context


# --------------------------------------------------------------------------- #
# 키워드 없음 / 텍스트 레이어 없음 안내
# --------------------------------------------------------------------------- #
def test_no_keyword_gives_note(sample_pdf: Path):
    report = pdf_service.search(sample_pdf, _criteria("존재하지않는단어"))
    assert report.total_matches == 0
    assert report.notes  # 안내 note 존재


def test_empty_text_layer_note(tmp_path: Path):
    """텍스트가 전혀 없는 PDF는 스캔/인코딩 가능성을 안내해야 한다."""
    doc = fitz.open()
    doc.new_page()  # 빈 페이지
    path = tmp_path / "blank.pdf"
    doc.save(path)
    doc.close()

    report = pdf_service.search(path, _criteria("무엇이든"))
    assert report.total_matches == 0
    assert any("스캔" in n for n in report.notes)


def test_exact_mode_adds_boundary_note(sample_pdf: Path):
    report = pdf_service.search(sample_pdf, _criteria("Secret", mode=SearchMode.EXACT))
    assert any("정확히 일치" in n for n in report.notes)


# --------------------------------------------------------------------------- #
# redaction 후 재검색 (실제 콘텐츠 제거 검증)
# --------------------------------------------------------------------------- #
def test_redaction_removes_text(sample_pdf: Path, tmp_path: Path):
    criteria = _criteria("CONFIDENTIAL", "Secret")
    req = EditRequest(criteria=criteria, pdf_action=PdfAction.REDACT)
    result = pdf_service.apply_edit(sample_pdf, req, tmp_path / "out")

    assert result.redactions_applied == 3  # page1 2건 + page2 1건
    assert Path(result.output_path).name == "sample_redacted.pdf"

    # 저장본을 직접 열어 텍스트가 실제로 사라졌는지 확인
    with fitz.open(result.output_path) as doc:
        full_text = "".join(page.get_text() for page in doc)
    assert "CONFIDENTIAL" not in full_text
    assert "Secret" not in full_text


def test_verify_after_redaction_is_clean(sample_pdf: Path, tmp_path: Path):
    criteria = _criteria("CONFIDENTIAL", "Secret")
    req = EditRequest(criteria=criteria, pdf_action=PdfAction.REDACT)
    result = pdf_service.apply_edit(sample_pdf, req, tmp_path / "out")
    verification = pdf_service.verify(Path(result.output_path), criteria)
    assert verification.clean is True


# --------------------------------------------------------------------------- #
# 원본 보존 / 미지원 형식
# --------------------------------------------------------------------------- #
def test_original_pdf_unchanged(sample_pdf: Path, tmp_path: Path):
    before = sample_pdf.read_bytes()
    req = EditRequest(criteria=_criteria("CONFIDENTIAL"), pdf_action=PdfAction.REDACT)
    result = pdf_service.apply_edit(sample_pdf, req, tmp_path / "out")
    assert sample_pdf.read_bytes() == before
    assert Path(result.output_path) != sample_pdf


def test_korean_pdf_search_and_redaction(tmp_path: Path):
    """한글 텍스트 PDF(내장 CJK 폰트)에서 검색·redaction·재검증이 모두 동작한다."""
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 100), "이 문서는 대외비 자료입니다", fontname="korea", fontsize=14)
    path = tmp_path / "kr.pdf"
    doc.save(path)
    doc.close()

    criteria = _criteria("대외비")
    report = pdf_service.search(path, criteria)
    assert report.total_matches == 1

    req = EditRequest(criteria=criteria, pdf_action=PdfAction.REDACT)
    result = pdf_service.apply_edit(path, req, tmp_path / "out")
    assert pdf_service.verify(Path(result.output_path), criteria).clean is True


def test_selected_redaction_only_selected(sample_pdf: Path, tmp_path: Path):
    """선택한 (페이지, 키워드)만 redaction하고 나머지는 그대로 둔다."""
    report = pdf_service.search(sample_pdf, _criteria("CONFIDENTIAL", "Secret"))
    chosen = [m for m in report.pdf_matches if m.keyword == "Secret"]  # page2 Secret만 선택

    req = EditRequest(criteria=_criteria("CONFIDENTIAL", "Secret"), pdf_action=PdfAction.REDACT)
    result = pdf_service.apply_edit(sample_pdf, req, tmp_path / "out", selected=chosen)

    with fitz.open(result.output_path) as doc:
        full_text = "".join(page.get_text() for page in doc)
    assert "Secret" not in full_text          # 선택됨 → 제거
    assert "CONFIDENTIAL" in full_text        # 미선택 → 유지
    assert result.redactions_applied == 1


def test_unsupported_extension_rejected(tmp_path: Path):
    bad = tmp_path / "doc.docx"
    bad.write_bytes(b"nope")
    with pytest.raises(ValueError, match="지원하지 않는"):
        pdf_service.search(bad, _criteria("x"))
