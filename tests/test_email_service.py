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
