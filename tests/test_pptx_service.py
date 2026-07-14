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


from document_redactor import pptx_service
from document_redactor.models import EditRequest


def test_search_finds_keywords_across_locations(make_pptx, tmp_path):
    p = make_pptx(
        tmp_path / "d.pptx",
        title="포스코 제목",
        body_lines=["대외비 본문"],
        table=[["포스코 셀", "일반"]],
        notes="노트에 대외비",
    )
    report = pptx_service.search(p, SearchCriteria(keywords=["포스코", "대외비"]))
    assert report.file_type == FileType.PPTX
    assert report.total_matches == 4  # 제목·표(포스코) + 본문·노트(대외비)
    locs = {m.location for m in report.pptx_matches}
    assert {"본문", "표", "노트"} <= locs


def test_apply_edit_removes_keywords_and_keeps_original(make_pptx, tmp_path):
    p = make_pptx(tmp_path / "d.pptx", title="포스코 제목", body_lines=["포스코 대외비"])
    out = tmp_path / "out"
    req = EditRequest(criteria=SearchCriteria(keywords=["포스코"]))
    result = pptx_service.apply_edit(p, req, out)

    produced = Path(result.output_path)
    assert produced.name == "d_edited.pptx"
    assert result.file_type == FileType.PPTX
    # 산출본에 포스코 없음, 원본은 그대로
    assert pptx_service.verify(produced, req.criteria).clean is True
    assert pptx_service.search(p, req.criteria).total_matches > 0


def test_apply_edit_handles_split_runs(make_pptx, tmp_path):
    # '포스'+'코 문서'로 쪼개진 런 → 런 단위 제거는 놓치지만 문단 폴백이 제거
    p = make_pptx(tmp_path / "s.pptx", title="제목", split_runs=["포스", "코 문서"])
    out = tmp_path / "out"
    req = EditRequest(criteria=SearchCriteria(keywords=["포스코"]))
    result = pptx_service.apply_edit(p, req, out)
    assert pptx_service.verify(Path(result.output_path), req.criteria).clean is True


def test_verify_detects_remaining(make_pptx, tmp_path):
    p = make_pptx(tmp_path / "d.pptx", title="포스코 제목")
    verification = pptx_service.verify(p, SearchCriteria(keywords=["포스코"]))
    assert verification.clean is False
    assert verification.remaining is not None
    assert verification.remaining.total_matches == 1
