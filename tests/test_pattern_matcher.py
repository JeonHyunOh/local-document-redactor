"""pattern_matcher 테스트 — 정형 개인정보 패턴 감지·제거·오탐 경계."""

from __future__ import annotations

from document_redactor import pattern_matcher as pm


def test_finds_always_on_types():
    text = "연락 010-1234-5678, 메일 a@b.com, 주민 900101-1234567, 카드 1234-5678-9012-3456"
    types = {t for t, _ in pm.find_patterns(text)}
    assert types == {"전화번호", "이메일", "주민등록번호", "신용카드"}


def test_account_is_opt_in():
    text = "계좌 123-45-678901"
    assert pm.find_patterns(text) == []  # 기본은 계좌 제외
    assert any(t == "계좌번호" for t, _ in pm.find_patterns(text, include_account=True))


def test_no_false_positive_on_plain_numbers():
    assert pm.find_patterns("제품 A-100, 일반숫자 123456") == []


def test_remove_patterns_strips_matches():
    out = pm.remove_patterns("메일 a@b.com 끝")
    assert "a@b.com" not in out and "메일" in out and "끝" in out


def test_remove_overlapping_prefers_longer():
    # 카드번호가 전체로 매칭·제거되어 숫자가 남지 않음
    out = pm.remove_patterns("카드 1234-5678-9012-3456 끝")
    assert not any(ch.isdigit() for ch in out)


def test_contains_pattern():
    assert pm.contains_pattern("a@b.com")
    assert not pm.contains_pattern("그냥 텍스트")
