"""PDF 검색·redaction 서비스.

검색(파일 무수정)과 편집(승인 후)을 분리한다. 규칙:
- 흰색 사각형 덮기 금지. search_for → add_redact_annot → apply_redactions로 실제 제거.
- 원본을 덮어쓰지 않고 ``_redacted`` 접미사 새 파일 생성 후 재검색으로 검증.
- 검색 0건이면 '없음'으로 단정하지 말고 스캔/인코딩 가능성을 notes로 안내.
- 대소문자 무시(case_sensitive=False)는 PyMuPDF 검색 기본 동작. 정확히 일치(EXACT)는
  단어 경계 판정이 어려워, PDF에서는 EXACT도 부분 문자열 검색 후 안내 note를 남긴다.
"""

from __future__ import annotations

from pathlib import Path

import fitz  # PyMuPDF

from .models import (
    EditRequest,
    EditResult,
    FileType,
    PdfMatch,
    SearchCriteria,
    SearchMode,
    SearchReport,
    VerificationResult,
)

_CONTEXT_RADIUS = 20  # 키워드 주변 문맥으로 보여줄 좌우 글자 수


def _resolve_pdf(path: Path) -> FileType:
    if path.suffix.lower() != ".pdf":
        raise ValueError(
            f"지원하지 않는 형식입니다: {path.suffix or '(확장자 없음)'}. PDF만 지원합니다."
        )
    return FileType.PDF


def _page_search(page: "fitz.Page", keyword: str) -> list:
    """페이지에서 키워드 좌표를 검색한다.

    PyMuPDF ``search_for``는 대소문자를 구분하지 않는다(기본 동작). MVP는 이를 그대로
    사용하고, 대소문자 구분이 요청되면 :func:`search`에서 note로 한계를 안내한다.
    """
    return page.search_for(keyword)


def _page_context(page_text: str, keyword: str, case_sensitive: bool) -> str | None:
    """페이지 텍스트에서 키워드 주변 문맥을 추출한다(가능한 경우)."""
    haystack = page_text if case_sensitive else page_text.casefold()
    needle = keyword if case_sensitive else keyword.casefold()
    idx = haystack.find(needle)
    if idx == -1:
        return None
    start = max(0, idx - _CONTEXT_RADIUS)
    end = min(len(page_text), idx + len(keyword) + _CONTEXT_RADIUS)
    snippet = page_text[start:end].replace("\n", " ").strip()
    return f"…{snippet}…"


def search(path: Path, criteria: SearchCriteria) -> SearchReport:
    """PDF를 수정하지 않고 키워드를 검사해 결과(좌표 포함)를 반환한다.

    텍스트 레이어가 없어 결과가 0건이면 스캔/인코딩 가능성을 notes로 안내한다.
    """
    _resolve_pdf(path)
    matches: list[PdfMatch] = []
    total_page_text_len = 0

    with fitz.open(path) as doc:
        for page_index, page in enumerate(doc):
            page_text = page.get_text()
            total_page_text_len += len(page_text.strip())
            for keyword in criteria.keywords:
                rects = _page_search(page, keyword)
                if not rects:
                    continue
                matches.append(
                    PdfMatch(
                        file_name=path.name,
                        page=page_index + 1,
                        keyword=keyword,
                        count=len(rects),
                        context=_page_context(page_text, keyword, criteria.case_sensitive),
                        rects=[tuple(r) for r in rects],
                    )
                )

    notes: list[str] = []
    if not matches:
        if total_page_text_len == 0:
            notes.append(
                "텍스트 레이어가 감지되지 않았습니다. 스캔 PDF이거나 글꼴/인코딩 문제로 "
                "검색이 불가능할 수 있습니다(OCR 미지원)."
            )
        else:
            notes.append("키워드를 찾지 못했습니다. 실제로 없거나, 글꼴/인코딩 문제일 수 있습니다.")
    if criteria.mode is SearchMode.EXACT:
        notes.append(
            "PDF 검색은 부분 문자열 기준입니다. '정확히 일치'는 단어 경계를 보장하지 않습니다."
        )
    if criteria.case_sensitive:
        notes.append(
            "PDF 검색은 영문 대소문자를 구분하지 않습니다. '대소문자 구분' 설정은 PDF에 적용되지 않습니다."
        )

    return SearchReport(
        file_name=path.name,
        file_type=FileType.PDF,
        criteria=criteria,
        pdf_matches=matches,
        notes=notes,
    )


def apply_edit(path: Path, request: EditRequest, output_dir: Path) -> EditResult:
    """승인된 redaction을 실행해 ``_redacted`` 접미사가 붙은 새 파일을 생성한다.

    페이지별로 모든 삭제 영역을 등록한 뒤 apply_redactions()로 실제 콘텐츠를 제거한다.
    원본을 덮어쓰지 않는다.
    """
    _resolve_pdf(path)
    criteria = request.criteria
    redactions_applied = 0
    log: list[str] = []

    with fitz.open(path) as doc:
        for page_index, page in enumerate(doc):
            page_redactions = 0
            for keyword in criteria.keywords:
                rects = _page_search(page, keyword)
                for rect in rects:
                    page.add_redact_annot(rect)
                    page_redactions += 1
            if page_redactions:
                page.apply_redactions()  # 페이지 등록 완료 후 일괄 적용
                redactions_applied += page_redactions
                log.append(f"p{page_index + 1}: {page_redactions}건 redaction 적용")

        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"{path.stem}_redacted{path.suffix}"
        doc.save(output_path, garbage=4, deflate=True)

    return EditResult(
        source_name=path.name,
        output_path=str(output_path),
        file_type=FileType.PDF,
        redactions_applied=redactions_applied,
        log=log,
    )


def verify(output_path: Path, criteria: SearchCriteria) -> VerificationResult:
    """redaction된 파일을 재검색해 키워드 잔존 여부를 검증한다."""
    report = search(output_path, criteria)
    return VerificationResult(
        output_path=str(output_path),
        clean=report.total_matches == 0,
        remaining=report if report.total_matches else None,
    )
