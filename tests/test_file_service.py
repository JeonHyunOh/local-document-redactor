"""file_service 단위 테스트 — 형식 판정·파일명 안전화·라우팅."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from document_redactor import email_service, file_service, pptx_service
from document_redactor.file_service import UnsupportedFileError
from document_redactor.models import FileType, SearchCriteria


def test_detect_supported_types():
    assert file_service.detect_file_type("a.xlsx") is FileType.XLSX
    assert file_service.detect_file_type("a.XLSM") is FileType.XLSM
    assert file_service.detect_file_type("report.pdf") is FileType.PDF


def test_detect_unsupported_raises():
    with pytest.raises(UnsupportedFileError, match="지원하지 않는"):
        file_service.detect_file_type("old.xls")


def test_safe_filename_strips_path_traversal():
    assert file_service.safe_filename("../../etc/passwd") == "passwd"
    assert file_service.safe_filename("/tmp/대외비 문서.xlsx") == "대외비_문서.xlsx"


def test_save_upload_writes_safe_name(tmp_path: Path):
    dest = tmp_path / "uploads"
    path = file_service.save_upload(b"data", "../weird name.pdf", dest)
    assert path.parent == dest
    assert path.name == "weird_name.pdf"
    assert path.read_bytes() == b"data"


def test_save_upload_rejects_unsupported(tmp_path: Path):
    with pytest.raises(UnsupportedFileError):
        file_service.save_upload(b"x", "old.hancom", tmp_path)


def test_routing_dispatches_to_excel(tmp_path: Path):
    wb = Workbook()
    wb.active["A1"] = "대외비"
    path = tmp_path / "r.xlsx"
    wb.save(path)
    report = file_service.search(path, SearchCriteria(keywords=["대외비"]))
    assert report.file_type is FileType.XLSX
    assert report.total_matches == 1


def test_detect_file_type_accepts_email():
    assert file_service.detect_file_type("a.msg") is FileType.MSG
    assert file_service.detect_file_type("a.eml") is FileType.EML


def test_detect_file_type_rejects_md_input():
    with pytest.raises(file_service.UnsupportedFileError):
        file_service.detect_file_type("a.md")


def test_service_for_routes_email_and_md():
    assert file_service._service_for(Path("a.msg")) is email_service
    assert file_service._service_for(Path("a.eml")) is email_service
    assert file_service._service_for(Path("out_redacted.md")) is email_service


def test_detect_and_route_pptx():
    assert file_service.detect_file_type("a.pptx") is FileType.PPTX
    assert file_service._service_for(Path("a.pptx")) is pptx_service


def test_detect_and_route_office_docs():
    from document_redactor import office_service

    assert file_service.detect_file_type("a.docx") is FileType.DOCX
    assert file_service.detect_file_type("a.hwp") is FileType.HWP
    for name in ("a.docx", "a.doc", "a.hwp", "a.hwpx"):
        assert file_service._service_for(Path(name)) is office_service


def test_remaining_rows_flattens_all_match_types():
    from document_redactor.models import (
        EmailMatch, ExcelMatch, PdfMatch, PptxMatch, SearchReport,
    )
    report = SearchReport(
        file_name="x", file_type=FileType.XLSX, criteria=SearchCriteria(keywords=["k"]),
        excel_matches=[ExcelMatch(file_name="x", sheet_name="Sheet1", cell="B3", row=3, keyword="대외비", original_value="대외비 문서")],
        pdf_matches=[PdfMatch(file_name="x", page=2, keyword="[전화번호]", count=1, context="010-1")],
        email_matches=[EmailMatch(file_name="x", field="본문", line=5, keyword="a@b.com", count=1, context="메일 a@b.com")],
        pptx_matches=[PptxMatch(file_name="x", slide=1, location="표", keyword="포스코", count=1, context="포스코 셀")],
    )
    rows = file_service.remaining_rows(report)
    items = {r["항목"] for r in rows}
    locs = {r["위치"] for r in rows}
    assert items == {"대외비", "[전화번호]", "a@b.com", "포스코"}
    assert "Sheet1!B3" in locs and "2페이지" in locs
    assert any("본문" in loc and "5" in loc for loc in locs)
    assert any("슬라이드1" in loc for loc in locs)
