"""batch_service 단위 테스트 — 폴더 스캔·배치 검색·배치 편집(구조 재현·오류 격리)."""

from __future__ import annotations

from pathlib import Path

import fitz
from openpyxl import Workbook, load_workbook

from document_redactor import batch_service, file_service
from document_redactor.models import (
    EditRequest,
    ExcelAction,
    SearchCriteria,
    VerificationResult,
)


def _xlsx(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.active["A1"] = value
    wb.save(path)


def _pdf(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), text, fontsize=14)
    doc.save(path)
    doc.close()


def _criteria(*kw: str) -> SearchCriteria:
    return SearchCriteria(keywords=list(kw))


def test_scan_folder_recursive_finds_supported(tmp_path: Path):
    _xlsx(tmp_path / "top.xlsx", "대외비")
    _pdf(tmp_path / "sub" / "deep.pdf", "대외비")
    (tmp_path / "note.txt").write_text("ignored")
    (tmp_path / "~$lock.xlsx").write_bytes(b"lock")  # Office 잠금 파일 제외

    found = batch_service.scan_folder(tmp_path, recursive=True)
    names = {p.name for p in found}
    assert names == {"top.xlsx", "deep.pdf"}


def test_scan_folder_non_recursive_skips_subdirs(tmp_path: Path):
    _xlsx(tmp_path / "top.xlsx", "x")
    _xlsx(tmp_path / "sub" / "nested.xlsx", "x")
    found = batch_service.scan_folder(tmp_path, recursive=False)
    assert {p.name for p in found} == {"top.xlsx"}


def test_batch_search_aggregates_and_isolates_errors(tmp_path: Path):
    _xlsx(tmp_path / "a.xlsx", "대외비 문서")
    _pdf(tmp_path / "b.pdf", "공개 자료")
    (tmp_path / "broken.pdf").write_bytes(b"%PDF-broken not a real pdf")

    items = batch_service.batch_search(tmp_path, _criteria("대외비"), recursive=True)
    by_name = {Path(i.path).name: i for i in items}
    assert by_name["a.xlsx"].matches == 1
    assert by_name["b.pdf"].matches == 0
    assert by_name["broken.pdf"].error is not None  # 오류 격리, 배치는 계속


def test_keyword_summary_and_match_details(tmp_path: Path):
    # PyMuPDF 기본 폰트는 한글 글리프를 못 실으므로 PDF 픽스처는 ASCII 키워드를 쓴다
    # (도구는 PDF를 생성하지 않고 기존 PDF를 검색만 하므로 실제 사용에는 무관).
    _xlsx(tmp_path / "a.xlsx", "대외비 및 내부검토용")  # 한 셀에 두 키워드
    _pdf(tmp_path / "sub" / "b.pdf", "SECRET SECRET report")  # 한 페이지 2회

    items = batch_service.batch_search(tmp_path, _criteria("대외비", "내부검토용", "SECRET"))

    summary = batch_service.keyword_summary(items)
    assert summary["대외비"] == 1
    assert summary["내부검토용"] == 1
    assert summary["SECRET"] == 2  # PDF 한 페이지에서 2회

    details = batch_service.match_details(items)
    # Excel 위치는 '시트!셀', PDF 위치는 'N페이지' 형식
    assert any(r["위치"] == "Sheet!A1" and r["키워드"] == "대외비" for r in details)
    assert any(r["위치"] == "1페이지" and r["키워드"] == "SECRET" for r in details)


def test_batch_search_reports_progress(tmp_path: Path):
    _xlsx(tmp_path / "a.xlsx", "x")
    _xlsx(tmp_path / "b.xlsx", "x")
    calls: list[tuple[int, int, str]] = []
    batch_service.batch_search(
        tmp_path, _criteria("x"), on_progress=lambda d, t, n: calls.append((d, t, n))
    )
    assert [c[0] for c in calls] == [1, 2]  # 완료 카운트 증가
    assert all(c[1] == 2 for c in calls)  # 전체 개수 일정


def test_batch_edit_mirrors_structure_and_verifies(tmp_path: Path):
    root = tmp_path / "src"
    _xlsx(root / "a.xlsx", "대외비 문서")
    _xlsx(root / "sub" / "b.xlsx", "대외비 메모")
    _xlsx(root / "clean.xlsx", "공개 자료")  # 매치 없음 → 사본 생성 안 함
    out = tmp_path / "out"

    req = EditRequest(criteria=_criteria("대외비"), excel_action=ExcelAction.CLEAR_CELL)
    items = batch_service.batch_edit(root, req, out, recursive=True)
    by_name = {Path(i.path).name: i for i in items}

    # 매치 없는 파일은 출력 없음
    assert by_name["clean.xlsx"].output_path is None

    # 하위 폴더 구조가 재현되어야 함
    edited_sub = out / "sub" / "b_edited.xlsx"
    assert edited_sub.exists()
    assert load_workbook(edited_sub).active["A1"].value is None

    # 편집된 파일은 모두 재검증 clean
    edited = [i for i in items if i.output_path]
    assert edited and all(i.clean for i in edited)

    # 원본은 그대로
    assert load_workbook(root / "a.xlsx").active["A1"].value == "대외비 문서"


def test_batch_edit_in_place_replaces_and_backs_up(tmp_path: Path):
    root = tmp_path / "src"
    _xlsx(root / "a.xlsx", "대외비 문서")
    _xlsx(root / "sub" / "b.xlsx", "대외비 메모")
    _xlsx(root / "clean.xlsx", "공개 자료")  # 매치 없음 → 교체·백업 안 함
    backup = tmp_path / "src_backup"

    req = EditRequest(criteria=_criteria("대외비"), excel_action=ExcelAction.CLEAR_CELL)
    items = batch_service.batch_edit_in_place(root, req, backup, recursive=True)
    by_name = {Path(i.path).name: i for i in items}

    # 원본이 제자리에서 편집됨
    assert load_workbook(root / "a.xlsx").active["A1"].value is None
    assert load_workbook(root / "sub" / "b.xlsx").active["A1"].value is None
    # 파일 경로/이름은 그대로 (접미사 없음)
    assert by_name["a.xlsx"].output_path == str(root / "a.xlsx")

    # 백업에 원본 내용이 구조 그대로 보존됨
    assert load_workbook(backup / "a.xlsx").active["A1"].value == "대외비 문서"
    assert load_workbook(backup / "sub" / "b.xlsx").active["A1"].value == "대외비 메모"

    # 매치 없는 파일은 교체·백업 안 함
    assert by_name["clean.xlsx"].output_path is None
    assert not (backup / "clean.xlsx").exists()


def test_batch_edit_in_place_keeps_original_when_verify_fails(tmp_path: Path, monkeypatch):
    """재검증 실패 시 원본을 덮어쓰지 않고 백업도 만들지 않는다(안전)."""
    root = tmp_path / "src"
    _xlsx(root / "a.xlsx", "대외비 문서")
    backup = tmp_path / "src_backup"

    # 재검증이 항상 실패한다고 가정
    monkeypatch.setattr(
        file_service,
        "verify",
        lambda out, crit: VerificationResult(output_path=str(out), clean=False),
    )

    req = EditRequest(criteria=_criteria("대외비"), excel_action=ExcelAction.CLEAR_CELL)
    items = batch_service.batch_edit_in_place(root, req, backup, recursive=True)

    # 원본 그대로, 교체 안 됨
    assert load_workbook(root / "a.xlsx").active["A1"].value == "대외비 문서"
    # 백업도 만들지 않음
    assert not backup.exists()
    # 결과에 실패 사유 기록
    assert items[0].error and "재검증 실패" in items[0].error
