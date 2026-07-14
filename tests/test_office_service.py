"""office_service 및 오피스 FileType 테스트 — 변환기는 monkeypatch로 대체."""

from __future__ import annotations

from pathlib import Path

import fitz

from document_redactor import doc_converter, office_service
from document_redactor.models import (
    EditRequest,
    FileType,
    SearchCriteria,
)


def test_office_filetypes():
    assert FileType.DOCX.value == "docx"
    assert FileType.DOC.value == "doc"
    assert FileType.HWP.value == "hwp"
    assert FileType.HWPX.value == "hwpx"


def _fake_pdf(tmp: Path, text: str) -> Path:
    p = tmp / "conv.pdf"
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), text)
    doc.save(str(p))
    doc.close()
    return p


def test_office_search_and_edit(tmp_path: Path, monkeypatch):
    # 변환기를 대체(실제 Office 없이). fake PDF는 fitz 폰트 한계를 피해 ASCII 사용.
    src = tmp_path / "보고서.docx"
    src.write_bytes(b"stub")
    monkeypatch.setattr(
        doc_converter, "convert_to_pdf", lambda p: _fake_pdf(tmp_path, "SECRET 010-1234-5678")
    )
    rep = office_service.search(src, SearchCriteria(keywords=["SECRET"]))
    assert rep.file_name == "보고서.docx"
    assert rep.total_matches >= 2  # 키워드 + 전화 패턴

    out = tmp_path / "out"
    res = office_service.apply_edit(
        src, EditRequest(criteria=SearchCriteria(keywords=["SECRET"])), out
    )
    assert Path(res.output_path).name == "보고서_redacted.pdf"
    assert office_service.verify(
        Path(res.output_path), SearchCriteria(keywords=["SECRET"])
    ).clean
