# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

로컬에서 동작하는 **엑셀·PDF 키워드 검사 및 삭제 도구**. 사용자가 파일과 키워드를 입력하면
문서에서 키워드 위치를 **검사(preview)** 하고, 사용자가 승인한 뒤에만 **삭제한 수정본**을 생성한다.
모든 처리는 로컬에서 수행하며 문서를 외부 서버·클라우드로 전송하지 않는다.

검색·편집은 전적으로 결정적(deterministic) 코드로 수행한다(openpyxl / PyMuPDF / 문자열 매칭).
LLM/AI는 사용하지 않는다 — 정확성·재현성·테스트 가능성을 최우선으로 한다.

## 개발 명령어

패키지 관리는 `pyproject.toml` (src 레이아웃 → editable 설치 필요).

```bash
# 가상환경 + 개발 설치 (테스트가 패키지를 import 하려면 -e 필요)
python3 -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"

# 앱 실행 (로컬 웹 UI)
streamlit run app.py

# 전체 테스트
pytest

# 단일 테스트 파일 / 단일 테스트
pytest tests/test_keyword_matcher.py
pytest tests/test_keyword_matcher.py::test_contains_korean -v

# 커버리지
pytest --cov=document_redactor
```

## 아키텍처: 반드시 지켜야 할 경계

**1. UI와 비즈니스 로직의 강한 분리**
`app.py`(Streamlit)에는 로직을 넣지 않는다. 파일 검사·편집·검증은 모두 `src/document_redactor/`
서비스 계층에 있어야 하며, `app.py`는 서비스 호출과 화면 표시만 담당한다. 서비스 함수는
Streamlit에 의존하지 않아야 한다(테스트에서 UI 없이 호출 가능해야 함).

**2. 검색(search)과 삭제(edit)는 별개의 단계 — 절대 합치지 않는다**
검색 단계는 파일을 절대 수정하지 않고 결과만 반환한다. 삭제는 사용자가 검색 결과를 보고
명시적으로 승인한 뒤에만 별도 함수로 실행한다. 이 두 단계가 한 함수에 섞이면 안 된다.

**3. 데이터 모델이 계층 간 계약이다**
`models.py`의 Pydantic 모델이 UI ↔ 서비스 사이의 유일한 데이터 계약이다.
검색 조건·삭제 방식 등 상태 문자열은 코드에 직접 쓰지 말고 `Enum`/`Literal`로 관리한다
(예: `SearchMode.CONTAINS/EXACT`, `ExcelAction.REMOVE_KEYWORD/CLEAR_CELL/DELETE_ROW`,
`PdfAction.REDACT`).

**4. 편집 결과는 항상 재검증한다**
수정본을 만든 직후 같은 검색 로직으로 다시 검사해 키워드 잔존 여부를 확인하고 `검증 결과`
모델로 반환한다. 검증 없이 "완료"로 처리하지 않는다.

### 표준 처리 흐름 (검색 → 미리보기 → 승인 → 편집 → 재검증)
업로드 → 키워드 입력 → 검색 조건 선택 → **검사(무수정)** → 표로 결과 표시 →
삭제 방식 선택 → **사용자 승인** → 원본 보존 상태로 수정본 생성 → **재검증** → 수정본+로그 다운로드.

## 안전 불변식 (위반 금지)

- **원본 파일을 절대 덮어쓰지 않는다.** 결과는 `_edited`(Excel) / `_redacted`(PDF) 등 접미사를 붙인 새 파일로 저장한다.
- 부분 실패한 수정본을 정상 결과처럼 제공하지 않는다. 실패 시 명확히 알리고 온전한 산출물만 넘긴다.
- 사용자의 명시적 승인 없이 실제 삭제를 실행하지 않는다.
- 예외를 `except Exception: pass`로 삼키지 않는다. 내부 오류 로그와 사용자 표시 메시지를 분리한다.
- 업로드 파일명을 그대로 신뢰하지 않고, 경로 처리는 `pathlib.Path`로만 한다. 임시 파일은 작업 후 정리한다.

## 파일 형식별 규칙

### Excel (openpyxl)
- 모든 워크시트의 셀을 검사한다.
- 삭제 방식: `키워드만 제거` / `셀 전체 비우기` / `행 전체 삭제`.
- **행 삭제**: 같은 행에서 여러 키워드가 나와도 중복 삭제하지 않는다. 행 번호가 밀리지 않도록
  **아래쪽 행부터(내림차순) 삭제**한다.
- `.xlsm`은 `load_workbook(..., keep_vba=True)`로 VBA 보존. 복잡한 매크로·외부 연결의 완전 보존은
  보장하지 않으며 UI에서 그 한계를 안내한다.
- **수식 셀**: 기본적으로 수식 문자열은 검사 대상에서 제외한다(계산 결과 손상 방지). 정책을 바꾸면
  코드와 문서에 명시한다.

### PDF (PyMuPDF / fitz)
- **단순 흰색 사각형 덮기 금지.** 반드시 redaction으로 실제 콘텐츠를 제거한다:
  `page.search_for()` → `page.add_redact_annot()` → 페이지별 등록 완료 후 `page.apply_redactions()`
  → 새 PDF 저장 → 저장본 재검색으로 검증.
- redaction 후 빈 공간이 남을 수 있음을 UI에서 안내한다.
- 검색 결과가 0건이면 "키워드 없음"으로 단정하지 말고 세 가능성을 함께 안내한다:
  (1) 실제 없음 (2) 스캔 PDF라 텍스트 레이어 없음 (3) 글꼴/인코딩 문제.

### 미지원 입력 처리
`.xls`, 암호화 파일, 스캔 전용 PDF, OCR 등은 임의로 처리하거나 실패시키지 말고 **명확한 안내
메시지**로 거부한다.

## 검색 로직

- 지원 모드: 포함(contains) / 정확히 일치(exact) / 영문 대소문자 구분 여부.
- 정규식은 MVP UI에 노출하지 않되, `keyword_matcher`는 모드 추가가 쉬운 구조로 설계한다.
- 키워드는 여러 개(한 줄에 하나) 입력. **정규화 시 빈 문자열·중복을 제거**한다.
- 한글 키워드가 1급 대상이므로 대소문자 규칙은 영문에만 적용됨에 유의한다.

## 폴더 배치 처리

`batch_service`가 폴더를 재귀 스캔해 여러 파일을 한 번에 처리한다.
- 파일별 오류는 격리한다(암호화·손상 파일은 건너뛰고 배치는 계속).
- `batch_edit`은 별도 출력 폴더에 원본 구조를 재현. `batch_edit_in_place`는 **교체 전 백업 →
  재검증 통과 시에만 제자리 교체**(실패 파일 원본은 유지). 진행률은 `on_progress` 콜백으로 전달.
- **확장자 무관 파일명 정리**: 제자리 모드의 파일명 키워드 정리는 지원 형식뿐 아니라 **모든 파일**에
  적용한다(내용 편집은 지원 형식만). 폴더명 정리는 기존대로 모든 하위 폴더 대상.
- **형식 완전 삭제(안전 불변식의 유일한 예외)**: `batch_edit_in_place(remove_suffixes=...)`는
  **opt-in + 명시적 승인** 시에만 `.dwg`/`.png`/`.nwd`(내용 정리가 불가능한 형식)를 **백업 없이
  완전 삭제**한다(`backup_root/_removed_log.txt`에 **확장자별 삭제 개수만** 기록 — 파일명은
  남기지 않는다(삭제 대상 파일명 자체가 민감정보일 수 있음), 복구 불가). 기본값은 비활성이며, 이
  경로 외에는 "원본 파일을 절대 삭제하지 않는다" 불변식을 그대로 지킨다.

## 작업 방식

- 기능 구현 전 현재 구조·요구사항을 먼저 확인하고, 작은 단위로 구현한다. 광범위한 일괄 리팩터링을 피한다.
- 파일 I/O 함수에는 타입 힌트를, public 함수에는 목적·주의사항 docstring을 단다.
- 새 기능 추가 시 관련 pytest 테스트를 함께 작성한다. 테스트용 파일은 저장소에 바이너리로 두지 말고
  테스트 실행 시 Python으로 생성한다.

## 구현 현황

- **모델·매칭**: `models.py`(Enum+Pydantic), `keyword_matcher.py`(정규화·매칭·제거).
- **Excel**: `excel_service.py` 검색·3종 삭제·재검증. 수식 셀·비문자열 제외, 행 내림차순 삭제.
- **PDF**: `pdf_service.py` redaction 검색·삭제·재검증. 텍스트 레이어 없음/EXACT/대소문자 안내 note.
- **배치**: `batch_service.py` 폴더 스캔·집계·별도폴더/제자리 편집·진행률.
- **UI**: `app.py`(Streamlit, 사이드바+탭) + `file_service.py`(유형 판정·안전 저장·라우팅). UI는 서비스만 호출.

향후 후보: 실제 `.xlsm`(VBA) 회귀 테스트, 검색 결과 CSV 내보내기, (요청 시) OCR로 스캔 PDF 지원.

## 서비스 계층 지도

`file_service`가 UI의 단일 진입점(라우터). `app.py` → `file_service` → (`excel_service`|`pdf_service`).
폴더 작업은 `app.py` → `batch_service` → `file_service`. `keyword_matcher`는 순수 로직(모든 계층이 공유).
