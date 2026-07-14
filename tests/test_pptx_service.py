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
