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


from document_redactor import email_service
from document_redactor.email_service import EmailContent


def _content(**kw):
    base = dict(subject="", sender="", to="", cc="", date="", body="", attachments=[])
    base.update(kw)
    return EmailContent(**base)


def test_render_markdown_omits_empty_fields():
    md = email_service.render_markdown(_content(subject="대외비 건", sender="홍길동"))
    assert "# 대외비 건" in md
    assert "- 보낸사람: 홍길동" in md
    assert "받는사람" not in md   # to 비어있으면 줄 생략


def test_render_markdown_lists_attachments_and_body():
    md = email_service.render_markdown(
        _content(subject="s", body="본문 내용", attachments=["a.xlsx", "b.pdf"])
    )
    assert "- 첨부파일: a.xlsx, b.pdf" in md
    assert "본문 내용" in md


def test_render_markdown_empty_body_placeholder():
    md = email_service.render_markdown(_content(subject="s"))
    assert "(본문 텍스트 없음)" in md


def test_find_matches_locates_keyword_in_body_with_field_and_count():
    content = _content(subject="공개", body="포스코 포스코 관련")
    matches = email_service._find_matches(content, SearchCriteria(keywords=["포스코"]), "a.eml")
    assert len(matches) == 1
    m = matches[0]
    assert m.field == "본문" and m.count == 2 and "포스코" in m.context


def test_find_matches_in_subject_and_attachment():
    content = _content(subject="포스코 보고", attachments=["포스코_첨부.xlsx"])
    kws = SearchCriteria(keywords=["포스코"])
    fields = {m.field for m in email_service._find_matches(content, kws, "a.eml")}
    assert "제목" in fields and "첨부" in fields


def test_find_matches_case_insensitive_default_and_sensitive():
    content = _content(subject="POSCO note")
    assert email_service._find_matches(content, SearchCriteria(keywords=["posco"]), "a")
    assert not email_service._find_matches(
        content, SearchCriteria(keywords=["posco"], case_sensitive=True), "a"
    )


def test_load_eml_normalizes_all_fields(make_eml, tmp_path):
    p = make_eml(
        tmp_path / "m.eml",
        subject="대외비 건",
        body="포스코 본문",
        sender="홍길동 <hong@example.com>",
        to="김철수 <kim@example.com>",
        cc="이영희 <lee@example.com>",
        attachments=["포스코_첨부.xlsx"],
    )
    content = email_service._load(p)
    assert content.subject == "대외비 건"
    assert "홍길동" in content.sender
    assert "김철수" in content.to
    assert "이영희" in content.cc
    assert "포스코 본문" in content.body
    assert content.attachments == ["포스코_첨부.xlsx"]


def test_load_msg_normalizes_core_fields(make_msg, tmp_path):
    p = make_msg(tmp_path / "m.msg", subject="대외비 건", body="포스코 본문", sender="홍길동")
    content = email_service._load(p)
    assert content.subject == "대외비 건"
    assert content.sender == "홍길동"
    assert "포스코 본문" in content.body


from document_redactor.models import EditRequest


def test_search_reports_email_matches(make_eml, tmp_path):
    p = make_eml(tmp_path / "m.eml", subject="포스코 보고", body="대외비 내용")
    report = email_service.search(p, SearchCriteria(keywords=["포스코", "대외비"]))
    assert report.file_type == FileType.EML
    assert report.total_matches == 2
    assert {m.keyword for m in report.email_matches} == {"포스코", "대외비"}


def test_apply_edit_produces_redacted_md_and_keeps_original(make_eml, tmp_path):
    p = make_eml(tmp_path / "m.eml", subject="포스코 보고", body="포스코 대외비 내용")
    out = tmp_path / "out"
    req = EditRequest(criteria=SearchCriteria(keywords=["포스코"]))
    result = email_service.apply_edit(p, req, out)

    produced = Path(result.output_path)
    assert produced.name == "m_redacted.md"
    text = produced.read_text(encoding="utf-8")
    assert "포스코" not in text
    assert "대외비 내용" in text          # 다른 텍스트는 보존
    assert result.file_type == FileType.MD
    assert result.redactions_applied == 2  # 제목 1 + 본문 1
    # 원본 .eml 미수정
    assert "포스코" in email_service._load(p).subject


def test_verify_clean_after_edit(make_eml, tmp_path):
    p = make_eml(tmp_path / "m.eml", subject="포스코", body="포스코")
    out = tmp_path / "out"
    req = EditRequest(criteria=SearchCriteria(keywords=["포스코"]))
    result = email_service.apply_edit(p, req, out)
    verification = email_service.verify(Path(result.output_path), req.criteria)
    assert verification.clean is True


def test_verify_detects_remaining(tmp_path):
    md = tmp_path / "x_redacted.md"
    md.write_text("# 포스코 남음\n\n본문", encoding="utf-8")
    verification = email_service.verify(md, SearchCriteria(keywords=["포스코"]))
    assert verification.clean is False
    assert verification.remaining is not None
    assert verification.remaining.total_matches == 1


def test_roundtrip_msg(make_msg, tmp_path):
    p = make_msg(tmp_path / "m.msg", subject="포스코 건", body="대외비 본문")
    out = tmp_path / "out"
    req = EditRequest(criteria=SearchCriteria(keywords=["포스코", "대외비"]))
    result = email_service.apply_edit(p, req, out)
    assert Path(result.output_path).name == "m_redacted.md"
    verification = email_service.verify(Path(result.output_path), req.criteria)
    assert verification.clean is True
