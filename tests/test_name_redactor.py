"""name_redactor 순수 로직 단위 테스트 — 이름에서 키워드 제거·정리·충돌 회피."""

from __future__ import annotations

from document_redactor.name_redactor import (
    name_contains_keyword,
    redact_filename,
    redact_segment,
    unique_name,
)


# --- name_contains_keyword ---
def test_name_contains_keyword_substring():
    assert name_contains_keyword("포스코_보고서.xlsx", ["포스코"], False)
    assert not name_contains_keyword("일반_보고서.xlsx", ["포스코"], False)


def test_name_contains_keyword_case_insensitive_default():
    assert name_contains_keyword("POSCO_report.pdf", ["posco"], False)


def test_name_contains_keyword_case_sensitive():
    assert not name_contains_keyword("POSCO_report.pdf", ["posco"], True)


# --- redact_filename: 확장자 보존, 위치 무관, 정리 ---
def test_redact_filename_middle_keeps_extension():
    assert redact_filename("2026_포스코_보고서.xlsx", ["포스코"], False) == "2026_보고서.xlsx"


def test_redact_filename_multiple_keywords():
    out = redact_filename("포스코_구형흑연폐수설비_설계.xlsx", ["포스코", "구형흑연폐수설비"], False)
    assert out == "설계.xlsx"


def test_redact_filename_whole_stem_falls_back():
    assert redact_filename("posco.pdf", ["posco"], False) == "redacted.pdf"


def test_redact_filename_preserves_double_extension_suffix_only():
    # 마지막 확장자만 suffix로 취급
    assert redact_filename("포스코.backup.xlsx", ["포스코"], False) == "backup.xlsx"


def test_redact_filename_case_insensitive():
    assert redact_filename("POSCO_a.xlsx", ["posco"], False) == "a.xlsx"


def test_redact_filename_no_keyword_unchanged():
    assert redact_filename("일반_보고서.xlsx", ["포스코"], False) == "일반_보고서.xlsx"


# --- redact_segment: 폴더명(확장자 없음) ---
def test_redact_segment_folder():
    assert redact_segment("01_포스코_구형화", ["포스코"], False) == "01_구형화"


def test_redact_segment_whole_falls_back():
    assert redact_segment("포스코", ["포스코"], False) == "redacted"


# --- unique_name: 충돌 시 접미사, 확장자 유지 ---
def test_unique_name_no_collision():
    assert unique_name("a.xlsx", set()) == "a.xlsx"


def test_unique_name_collision_suffixes_before_extension():
    taken = {"a.xlsx", "a_1.xlsx"}
    assert unique_name("a.xlsx", taken) == "a_2.xlsx"


def test_unique_name_collision_no_extension():
    assert unique_name("folder", {"folder"}) == "folder_1"
