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
        file_service.save_upload(b"x", "macro.docx", tmp_path)


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
