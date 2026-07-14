"""doc_converter 테스트 — 실제 변환은 Word/한글 COM 가용 시에만 실행(그 외 skip)."""

from __future__ import annotations

from pathlib import Path

import pytest

from document_redactor import doc_converter


def _word_available() -> bool:
    try:
        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()
        try:
            w = win32.Dispatch("Word.Application")
            w.Quit()
            return True
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        return False


def test_unsupported_suffix_raises(tmp_path: Path):
    p = tmp_path / "a.txt"
    p.write_text("x", encoding="utf-8")
    with pytest.raises(doc_converter.ConversionError):
        doc_converter.convert_to_pdf(p)


@pytest.mark.skipif(not _word_available(), reason="Word COM 미가용")
def test_word_docx_to_pdf(tmp_path: Path):
    import pythoncom
    import win32com.client as win32

    docx = tmp_path / "t.docx"
    pythoncom.CoInitialize()
    w = win32.Dispatch("Word.Application")
    w.Visible = False
    d = w.Documents.Add()
    d.Content.Text = "대외비 문서"
    d.SaveAs(str(docx), 16)
    d.Close(False)
    w.Quit()
    pythoncom.CoUninitialize()

    pdf = doc_converter.convert_to_pdf(docx)
    import fitz

    assert "대외비" in fitz.open(str(pdf))[0].get_text()
