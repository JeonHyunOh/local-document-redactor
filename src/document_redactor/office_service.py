"""오피스 문서(docx/doc/hwp/hwpx)를 PDF로 변환한 뒤 pdf_service로 검사·redaction한다.

변환은 doc_converter(COM)에 격리되어 있고, 이 서비스는 그 결과 PDF를 기존 pdf_service로
검사·삭제·재검증한다. 원본 문서는 편집하지 않으며, 정리 결과는 <원본stem>_redacted.pdf로
산출한다. doc_converter.convert_to_pdf를 대체(monkeypatch)하면 실제 Office 없이 테스트할 수 있다.
"""

from __future__ import annotations

from pathlib import Path

from . import doc_converter, pdf_service
from .models import (
    EditRequest,
    EditResult,
    SearchCriteria,
    SearchReport,
    VerificationResult,
)

_RENDER_NOTE = "원본 문서를 PDF로 변환해 검사·정리했습니다(결과는 PDF)."


def search(path: Path, criteria: SearchCriteria) -> SearchReport:
    """오피스 문서를 PDF로 변환해 키워드·패턴을 검사한다(원본 무수정)."""
    path = Path(path)
    pdf = doc_converter.convert_to_pdf(path)
    report = pdf_service.search(pdf, criteria)
    report.file_name = path.name
    if _RENDER_NOTE not in report.notes:
        report.notes.append(_RENDER_NOTE)
    return report


def apply_edit(path: Path, request: EditRequest, output_dir: Path, selected=None) -> EditResult:
    """변환된 PDF에서 키워드·패턴을 제거해 <원본stem>_redacted.pdf로 저장한다."""
    path = Path(path)
    pdf = doc_converter.convert_to_pdf(path)
    edit = pdf_service.apply_edit(pdf, request, output_dir)
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
    """산출물은 .pdf이므로 pdf_service.verify로 위임한다."""
    return pdf_service.verify(Path(output_path), criteria)
