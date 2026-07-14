# 패턴 자동 삭제 + ZIP 처리 — 설계

- 작성일: 2026-07-14
- 상태: 승인됨 (구현 대기)
- 대상 레포: `local-document-redactor`

## 배경 / 목표

두 가지 기능을 추가한다.

1. **패턴 자동 삭제**: 사용자가 입력한 키워드와 **무관하게**, 전화번호·이메일·주민등록번호·
   신용카드번호 같은 **정형 개인정보 패턴**을 항상 자동 감지·제거한다. (계좌번호는 오탐 위험이 커
   **opt-in 체크박스**로 분리.)
2. **ZIP 처리(제자리 모드 전용)**: 폴더 안의 `.zip`을 압축 해제해 내부 파일을 정상 파이프라인
   (키워드·패턴 삭제, 파일명 정리, 형식 제거)으로 처리하고, 원본 `.zip`은 백업 후 삭제한다.

## 핵심 결정 (사용자 승인)

- 패턴: **이메일·주민등록번호·신용카드·전화번호는 항상 자동 적용**. **계좌번호는 opt-in**(기본 꺼짐,
  오탐 위험 안내). 모든 콘텐츠 형식(excel·pdf·pptx·email)과 모든 모드에 적용.
- ZIP: **제자리 모드에서만**. 중첩 zip은 `.zip`이 없어질 때까지 반복 해제. 원본 zip은 백업 후 삭제.
- 오탐 위험이 있는 패턴(전화·계좌)은 정규식을 보수적으로 잡고 UI에 한계를 안내한다.

## 비목표
- 패턴의 문맥 기반 검증(예: 유효한 카드 Luhn 체크)까지는 하지 않는다(단순 형식 매칭).
- 암호화 zip 해제(비밀번호 입력)·7z/rar 등 다른 압축 형식은 범위 밖(명확히 거부·격리).
- 출력본 모드의 zip 처리(요구 범위 밖, 후속).

## A. 패턴 자동 삭제

### 신규 순수 모듈 `src/document_redactor/pattern_matcher.py`

```
_PATTERNS: dict[str, re.Pattern]   # 항상 적용 유형
_ACCOUNT_PATTERN: re.Pattern        # opt-in 계좌번호

find_patterns(text: str, include_account: bool = False) -> list[PatternHit]
    # (유형, 매칭 문자열, span) 목록. 구체적 패턴(RRN·카드) 우선, 겹침은 긴 매치 우선.

remove_patterns(text: str, include_account: bool = False) -> str
    # 매칭 구간을 모두 제거한 문자열 반환(키워드 remove와 동일하게 부분 제거).

contains_pattern(text: str, include_account: bool = False) -> bool
```

**패턴(초안, 보수적):**
- 이메일: `[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}`
- 주민등록번호: `\b\d{6}-\d{7}\b`
- 신용카드: `\b\d{4}[- ]\d{4}[- ]\d{4}[- ]\d{4}\b`(구분자 있는 형태)
- 전화번호: `\b(?:\+?82[- ]?)?0\d{1,2}[- ]?\d{3,4}[- ]?\d{4}\b`
- 계좌번호(opt-in): `\b\d{2,6}-\d{2,6}-\d{2,6}(?:-\d{1,6})?\b`(느슨 → 오탐 안내)

겹침 처리: 모든 유형의 매치를 모아 **span 기준 비겹침 병합(긴 매치 우선)** 후 제거.

### 데이터 모델 (`models.py`)
- `SearchCriteria`에 `redact_account_numbers: bool = False` 추가(패턴 opt-in 계좌).
- 항상 적용 패턴은 별도 플래그 없이 서비스가 항상 수행한다.
- 패턴 매치 보고: 각 형식의 기존 match 모델을 재사용하되, `keyword` 필드에 **유형 라벨**(예:
  `[전화번호]`)을 넣어 키워드 매치와 구분한다(매칭된 실제 값은 context/original_value에 담김).

### 서비스 통합 (search / edit / verify 모두)
모든 콘텐츠 서비스가 **키워드 처리와 나란히 패턴 처리**를 수행한다.

- **excel_service**: 셀 텍스트에서 `find_patterns`로 매치 추가 검출. 편집은 `remove_keywords` 후
  `remove_patterns`도 적용(REMOVE_KEYWORD). CLEAR_CELL/DELETE_ROW은 패턴 매치 셀도 대상.
- **pptx_service / email_service**: 문단/렌더 텍스트에 `remove_patterns` 추가. 검색·재검증도 패턴 포함.
- **pdf_service**: `get_text()`로 페이지 텍스트 추출 → `find_patterns`로 매칭 **문자열**을 얻고,
  그 문자열을 `search_for`로 위치를 잡아 redaction. (레이아웃상 공백 차이로 못 찾는 경우는 note로 안내.)
- **verify**: 재검색에 패턴을 포함해 키워드·패턴 잔존을 모두 확인한다.

`include_account`는 `SearchCriteria.redact_account_numbers`에서 전달.

### UI (`app.py`)
- 사이드바에 **"개인정보 패턴 자동 삭제"** 안내(항상 적용됨) + **"계좌번호도 삭제(오탐 주의)"** 체크박스.
- 검색 결과 표에 패턴 매치가 `[전화번호]` 등 유형 라벨로 표시됨.
- 계좌번호 체크 시 "일반 숫자열이 함께 지워질 수 있음" 경고.

## B. ZIP 처리 (제자리 모드 전용)

### `batch_service.batch_edit_in_place`에 새 **Phase Z(가장 먼저)** 추가

시그니처에 영향 없음(zip 처리는 항상 수행하거나, 안전하게 기본 수행). 순서:

1. **Phase Z — 압축 해제**: `scan_all_files`로 `.zip`을 찾아, 각 zip을 같은 디렉터리의
   `<zip stem>/` 폴더로 해제(폴더명 충돌 시 `_1` 접미사). 해제 성공 시 **원본 zip을
   `backup_root`에 백업한 뒤 삭제**. `.zip`이 더 없을 때까지 **반복**(중첩 zip 대응).
   - 암호/손상/미지원 → 격리(zip 유지, 결과에 오류 기록, 배치 계속).
   - 결과 항목: `BatchEditItem(note="압축 해제 후 원본 zip 삭제")`.
2. 이후 기존 **Phase 0(형식 제거) → Phase 1(내용 편집: 키워드+패턴) → Phase 2(이름 정리)** 가
   **해제된 파일들까지 포함**해 처리한다(압축 해제를 먼저 했으므로 자동 포함).

### 안전
- 원본 zip은 삭제 전 `backup_root`에 백업(복구 가능).
- 압축 해제 시 **zip slip 방지**: 각 엔트리 경로가 대상 폴더 밖(`..`)을 벗어나면 거부.
- 경로는 `pathlib`, 해제는 표준 `zipfile`.

## 재검증 (안전 불변식)
- 키워드·패턴 모두 재검증에 포함. 편집 후 잔존 여부를 함께 확인한다.
- zip: 해제·삭제는 결과/로그로 근거 제공(원본은 백업).

## 테스트 (pytest, 픽스처는 실행 시 생성)
- `tests/test_pattern_matcher.py`: 각 유형 매칭·제거, 겹침 병합, 계좌 opt-in on/off, 오탐 경계
  (예: 일반 6자리 숫자는 RRN 아님, 구분자 없는 16자리는 카드로 안 잡음 등).
- 각 서비스 테스트 확장: 패턴이 키워드와 함께 검색·제거·재검증됨(excel·pdf·pptx·email).
- `tests/test_batch_service.py`: zip 해제→내부 파일 키워드·패턴 정리→원본 zip 삭제·백업,
  중첩 zip 반복 해제, zip slip 거부, 손상 zip 격리.
- `tests/test_app_smoke.py`: import 무결성.

## 미해결/후속 (범위 밖)
- 출력본 모드 zip 처리.
- 암호 zip(비밀번호), 7z/rar.
- 카드 Luhn 검증 등 패턴 정밀도 향상.
