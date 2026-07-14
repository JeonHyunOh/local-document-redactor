# 형식 제거·확장자 무관 파일명 정리·pptx 지원 — 설계

- 작성일: 2026-07-14
- 상태: 승인됨 (구현 대기)
- 대상 레포: `local-document-redactor`
- 선행 기능: 파일명·폴더명 정리([2026-07-13 …](2026-07-13-filename-folder-keyword-redaction-design.md)),
  이메일 → md([2026-07-14 …](2026-07-14-email-to-markdown-redaction-design.md)) — `name_redactor`·서비스 계약 재사용.

## 배경 / 문제

세 가지 사용자 요구를 한 번에 반영한다(모두 배치 처리 강화, pptx는 단일 파일도 포함).

1. `.dwg`/`.png`/`.nwd`(CAD·이미지·3D)는 내용 정리가 불가능하다. 정리 대상 폴더에서
   **파일 자체를 완전 삭제**하고 싶다(제자리 모드).
2. 파일명 키워드 정리가 **지원 형식(xlsx/xlsm/pdf/msg/eml) 파일에만** 적용되고 있다
   (`scan_folder`가 지원 확장자만 반환). **확장자 무관 모든 파일**의 파일명을 정리해야 한다(제자리 모드).
3. 내용 키워드 삭제가 xlsx/pdf/eml/msg만 지원한다. **`.pptx`** 도 대상에 추가한다.

## 목표 / 결정 (사용자 승인 완료)

1. **형식 제거**: 확장자 `{.dwg, .png, .nwd}` 파일을 **제자리 모드에서 완전 삭제**(백업 없음).
   **opt-in**(기본 꺼짐) + 승인 게이트. `_removed_log.txt`에 삭제 목록 기록(복구 불가 안내).
2. **확장자 무관 파일명 정리**: 제자리 모드에서 **모든 파일**(확장자 무관)의 파일명 키워드를 정리.
3. **pptx 내용 삭제**: `pptx_service`(python-pptx)로 슬라이드·도형·표·노트 텍스트에서 키워드 제거.
   단일 파일 + 폴더(출력본·제자리) 모두.

**우선순위(제자리 모드)**: **제거(.dwg/.png/.nwd) > 내용 편집 + 파일명 정리.**
제거 대상 파일은 내용 편집·파일명 정리 대상이 아니다(먼저 삭제).

## 비목표

- `.dwg`/`.png`/`.nwd`의 내용 파싱·정리(포맷 특성상 불가).
- 형식 제거·확장자 무관 파일명 정리의 **출력본 모드** 확장(요구는 제자리 모드로 한정).
  - 출력본 모드는 원래 처리 파일만 새 폴더에 쓰므로 `.dwg/.png/.nwd`는 자연히 제외된다.
- pptx의 이미지·차트·SmartArt 내부 텍스트, 도형 그룹의 완전 재귀(그룹은 1차 지원, 한계 안내).

## 안전 불변식 변경 (CLAUDE.md 갱신 포함)

형식 제거는 **"원본 파일을 절대 삭제하지 않는다"** 불변식에 예외를 만든다. 이를 문서와 코드에
일치시킨다:
- CLAUDE.md의 배치 처리 절에 "**opt-in + 명시적 승인** 시에만 `.dwg/.png/.nwd`를 완전 삭제할 수
  있다(제자리 모드, 백업 없음, `_removed_log.txt` 기록)"를 명시한다.
- 기본값은 비활성. 사용자가 체크박스로 켜고 승인해야만 삭제된다.

## 아키텍처

기존 계층 경계 유지. UI에는 로직을 넣지 않는다.

### 신규 모듈: `src/document_redactor/pptx_service.py`

`excel_service`/`pdf_service`와 동일 계약(`search`/`apply_edit`/`verify`). python-pptx 의존을
이 모듈에 격리한다. 쪼개진 런(run) 대응을 위해 **검색·재검증은 문단 단위**, 제거는 런 단위 후
남으면 **문단 단위 폴백**.

```
_iter_text_frames(prs) -> Iterator[(slide_index, location, text_frame)]
    # location: "본문"/"표"/"노트". 도형·표·노트를 순회.

search(path, criteria) -> SearchReport
    # 각 문단의 런 텍스트를 합쳐 키워드 판정 → PptxMatch(slide, location, keyword, count, context).

apply_edit(path, request, output_dir, selected=None) -> EditResult
    # 런 단위 remove_keywords 후, 문단에 키워드가 남으면(런 분할) 첫 런에 정리된 문단 텍스트를
    # 넣고 나머지 런을 비워 확실히 제거. <stem>_edited.pptx 저장. selected 무시(전체 제거).

verify(output_path, criteria) -> VerificationResult
    # 저장본을 문단 단위로 재검색. 키워드 잔존 없으면 clean=True.
```

### 데이터 모델 (`models.py`)
- `FileType`에 `PPTX = "pptx"` 추가.
- `PptxMatch(file_name, slide: int, location: str, keyword, count, context)`.
- `SearchReport`에 `pptx_matches: list[PptxMatch]` 추가 + `total_matches` 합산.
- 형식 제거 결과: `BatchEditItem`에 담지 않고, in-place가 반환하는 결과와 별도로
  `removed: list[str]`(삭제된 상대경로) 및 로그로 표현한다. → 반환 타입을 바꾸지 않기 위해
  `batch_edit_in_place`는 기존 `list[BatchEditItem]`을 유지하고, 삭제된 파일은 `BatchEditItem`에
  `note="완전 삭제됨(.dwg/.png/.nwd)"` + `error=None` + `output_path=None`으로 표기한다.

### 라우팅 (`file_service.py`)
- `_EXTENSION_MAP`에 `.pptx → FileType.PPTX`. 안내 문구에 `.pptx` 추가.
- `_service_for`: `.pptx` → `pptx_service`.

### 배치 (`batch_service.py`)
- `_SUPPORTED_SUFFIXES`에 `.pptx` 추가(내용 처리 대상).
- 신규 상수 `_REMOVAL_SUFFIXES = {".dwg", ".png", ".nwd"}`.
- 신규 헬퍼 `scan_all_files(root, recursive)` — 숨김·`~$` 제외한 **모든 파일**(확장자 무관) 반환.
- `batch_edit`(출력본): `.pptx`가 지원 형식이 되어 자동으로 편집 대상에 포함(기존 로직 재사용).
  형식 제거·확장자 무관 rename은 출력본 모드에서 하지 않는다(요구 범위).
- `batch_edit_in_place`(제자리): 아래 단계로 확장. 신규 파라미터
  `remove_suffixes: set[str] | None = None`(형식 제거 opt-in; None이면 제거 안 함).

  **Phase 0 — 형식 제거(opt-in)**: `remove_suffixes`가 주어지면, 해당 확장자 파일을 `pathlib`로
  **완전 삭제**(`Path.unlink`). 백업하지 않는다. 삭제 목록을 모아 `_removed_log.txt`에 기록하고,
  각 파일을 `BatchEditItem(note="완전 삭제됨 …")`로 결과에 남긴다. 이후 단계는 삭제된 파일을 보지 않는다.

  **Phase 1 — 내용 편집(기존)**: `scan_folder`(지원 형식, 이제 pptx 포함) 대상. 이메일은 기존대로
  내용 미처리+note. 재검증 통과 시에만 백업 후 제자리 교체.

  **Phase 2a — 파일명 rename(확장자 무관으로 확장)**: `scan_all_files`로 **현재 폴더의 모든 파일**을
  순회. 파일명에 키워드가 있으면 확장자 무관 rename. 기존 `by_path` 항목이 있으면 갱신, 없으면
  (미지원 파일) 새 `BatchEditItem(renamed_to=…)` 추가. rename 전 백업 보장(기존 규칙 유지).
  제거 대상은 Phase 0에서 이미 삭제되어 여기 없음.

  **Phase 2b — 폴더 rename(기존)**: bottom-up, 루트 제외.

  `_rename_log.txt`(rename)와 `_removed_log.txt`(삭제)는 별도로 기록.

### UI (`app.py`)
- 업로더 `type`에 `"pptx"` 추가. 단일 파일 pptx는 excel처럼 매치 미리보기(슬라이드·위치·키워드·문맥)
  후 승인 → `<stem>_edited.pptx` 다운로드(삭제 방식 선택 없음, 항상 키워드 제거).
- 폴더 제자리 모드에만: **"CAD·이미지·3D 파일(.dwg/.png/.nwd) 완전 삭제"** 체크박스(기본 꺼짐).
  켜면 승인 문구에 삭제 대상 개수 표시. 결과에 삭제 목록·`_removed_log.txt` 위치 안내.
- 폴더 결과표: `.pptx` 산출물 표기, 삭제된 파일 note 표기.

## 재검증 (안전 불변식 유지)
- pptx: 저장본 문단 단위 재검색으로 잔존 확인.
- 형식 제거: 삭제 후 해당 경로 부재 확인(로그로 근거 제공).
- 확장자 무관 rename: 최종 파일명에 키워드 잔존 여부 확인(기존 이름 재검증 확장).

## 테스트 (pytest, 파일은 실행 시 생성)

- `tests/test_pptx_service.py`: python-pptx로 픽스처 생성(텍스트박스·표·노트). search(슬라이드/위치/
  카운트), apply_edit(`_edited.pptx`·키워드 제거·원본 보존), verify(clean/남음), 런 분할 케이스
  (한 문단을 여러 런으로 만들어 문단 단위 제거 확인).
- `tests/test_file_service.py`: `.pptx` 업로드 허용·라우팅.
- `tests/test_batch_service.py`:
  - `scan_all_files`가 확장자 무관 모든 파일 반환(숨김·`~$` 제외).
  - 제자리 Phase 0: `remove_suffixes` 지정 시 `.dwg/.png/.nwd` 완전 삭제 + `_removed_log.txt`.
  - 제자리 Phase 2a: 미지원 확장자(.txt 등) 파일명에 키워드 있으면 rename됨.
  - 제거 opt-in 꺼짐(기본): 삭제 안 함.
  - pptx 내용 배치 편집(출력본·제자리).
- `tests/test_app_smoke.py`: import 무결성.

## 미해결/후속 (범위 밖)
- 출력본 모드의 형식 제거·확장자 무관 rename(요구 범위 밖).
- pptx 그룹 도형 깊은 재귀·차트/SmartArt 텍스트.
- `.dwg/.png/.nwd` 내용 기반 처리(불가).
