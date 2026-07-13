# .msg 분석 → .md 산출(키워드 제거) — 설계

- 작성일: 2026-07-14
- 상태: 승인됨 (구현 대기)
- 대상 레포: `local-document-redactor`
- 선행 기능: 파일명·폴더명 키워드 제거([2026-07-13-filename-folder-keyword-redaction-design.md](2026-07-13-filename-folder-keyword-redaction-design.md)) — `name_redactor` 재사용.

## 배경 / 문제

현재 도구는 `.xlsx` / `.xlsm` / `.pdf`만 지원한다(`file_service._EXTENSION_MAP`). Outlook 메시지
파일 `.msg`는 업로드 시 `UnsupportedFileError`로 거부된다. 그러나 `.msg`에는 제목·보낸사람·
본문 등에 민감 키워드가 그대로 들어 있어 검사·정리 대상이 되어야 한다.

`.msg`는 OLE 복합 파일이다. **읽기(파싱)는 `extract-msg`로 안정적이지만, `.msg` 구조를 편집해
다시 저장하는 검증된 방법은 없다.** 이 프로젝트의 안전 불변식(“수정본 생성 후 재검증”,
“부분 실패한 수정본을 정상처럼 제공 금지”)과 충돌하므로, `.msg` 자체를 편집하지 않는다.

## 목표

`.msg`를 분석해 제목·보낸사람·받는사람·참조·날짜·첨부목록·본문을 **Markdown 텍스트로 렌더링**
하고, 그 텍스트에서 등록 키워드를 제거한 **`.md` 파일을 산출**한다. 원본 `.msg`는 절대 편집하지
않는다(별도 새 파일 생성).

## 비목표 (범위 밖)

- `.msg` 자체(OLE 스트림) 편집/재저장.
- 첨부파일 **내용** 추출·검사·삭제. (첨부는 **이름만** 목록으로 표기)
- `.msg`의 제자리 교체(형식이 `.md`로 바뀌므로 “원본 덮어쓰기”가 성립하지 않음).
- 본문 서식·이미지·HTML 레이아웃 보존. 본문은 평문 텍스트로만 다룬다.

## 핵심 결정 (사용자 승인 완료)

1. **처리 방식**: `.msg` 분석 → 키워드 제거한 `.md` 산출. 원본 `.msg` 보존.
2. **첨부파일**: `.md`에 **이름만 목록**으로 표기. 이름에 키워드가 있으면 그 이름에서도 제거.
   첨부 내용은 추출·처리하지 않는다.
3. **모드**: **단일 파일** + **폴더 출력본(batch_edit)**. 제자리 교체 모드는 `.msg` 내용을
   처리하지 않는다(아래 §배치 참조).
4. **라이브러리**: `extract-msg`(읽기 전용) 신규 의존성 추가.
5. **부분 선택 미지원**: `.msg` 편집은 **모든 키워드 occurrence 제거**로 고정한다(Excel/PDF
   단일 파일 모드의 항목별 체크 선택은 `.msg`에 적용하지 않는다 — 텍스트 산출 특성상 단순화).

## 아키텍처

기존 계층 경계를 유지한다. `excel_service`/`pdf_service`와 동일한 서비스 계약
(`search` / `apply_edit` / `verify`)을 구현하는 **새 서비스 `msg_service.py`** 를 추가하고,
`file_service`가 `.msg`를 이 서비스로 라우팅한다. UI(`app.py`)에는 로직을 넣지 않는다.

### extract-msg 의존성 격리 (테스트 가능성의 핵심)

`extract-msg`에 직접 의존하는 코드는 **얇은 어댑터 한 곳**에만 둔다. 나머지 로직(렌더·검색·
제거·검증)은 정규화된 순수 데이터 구조 위에서 동작해 `.msg` 파일 없이 단위 테스트할 수 있다.

```python
@dataclass(frozen=True)
class MsgContent:
    subject: str
    sender: str
    to: str
    cc: str
    date: str
    body: str
    attachments: list[str]   # 첨부 파일명 목록(내용 아님)

def _load_message(path: Path) -> MsgContent:
    """extract_msg로 .msg를 파싱해 MsgContent로 정규화한다. (유일한 extract_msg 의존 지점)"""
```

### 신규 모듈: `src/document_redactor/msg_service.py`

```
render_markdown(content: MsgContent) -> str
    # MsgContent를 정해진 Markdown 틀로 렌더링(결정적).

_labeled_lines(content: MsgContent) -> list[tuple[str, str]]
    # (field 라벨, 줄 텍스트) 목록. 검색이 매치의 field/줄번호를 붙이는 데 사용.
    # render_markdown 출력과 줄 순서가 일치한다.

search(path: Path, criteria: SearchCriteria) -> SearchReport
    # _load_message → _labeled_lines 각 줄에서 keyword_matcher로 매치 → MsgMatch 목록.
    # 파일 무수정. file_type=MSG.

apply_edit(path, request, output_dir, selected=None) -> EditResult
    # render_markdown 결과에서 keyword_matcher.remove_keywords로 키워드 제거 →
    # <safe_stem>_redacted.md 저장. selected는 무시(항상 전체 제거, §핵심결정 5).

verify(output_path: Path, criteria: SearchCriteria) -> VerificationResult
    # 산출된 .md 텍스트를 줄 단위로 재검색해 키워드 잔존 여부 확인.
```

### 렌더 틀 (결정적)

```
# <제목>

- 보낸사람: <sender>
- 받는사람: <to>
- 참조: <cc>
- 날짜: <date>
- 첨부파일: a.xlsx, b.pdf

---

<본문 평문>
```

- 빈 필드(예: cc 없음)는 해당 줄을 생략한다.
- 첨부가 없으면 “첨부파일” 줄 생략.
- 본문 평문(`body`)이 없고 HTML만 있으면 텍스트로 정제하거나 “(본문 텍스트 없음)” 표기.

### 키워드 제거 정책

- 편집은 렌더된 **전체 Markdown 텍스트**에 `keyword_matcher.remove_keywords`를 적용한다.
  이렇게 하면 제목·헤더·본문·첨부 이름에서 키워드가 **일관되게 부분문자열로 제거**된다
  (내용 검색 모드가 EXACT여도 제거는 항상 부분문자열 — 기존 서비스의 REMOVE 동작과 일치).
- `.msg` 내용 검색(미리보기)은 PDF처럼 **텍스트 기반**이다. CONTAINS는 부분문자열, EXACT는
  줄 전체 일치. 실사용은 CONTAINS가 기본.

## 데이터 모델 (`models.py`)

- `FileType`에 `MSG = "msg"` 추가.
- 신규 `MsgMatch`:
  ```python
  class MsgMatch(BaseModel):
      file_name: str
      field: str      # "제목" / "보낸사람" / "받는사람" / "참조" / "날짜" / "첨부" / "본문"
      line: int       # 렌더된 .md에서의 1-기반 줄 번호
      keyword: str
      count: int      # 해당 줄에서 발견된 횟수
      context: str    # 해당 줄 텍스트
  ```
- `SearchReport`에 `msg_matches: list[MsgMatch] = Field(default_factory=list)` 추가.
- `SearchReport.total_matches`에 `sum(m.count for m in msg_matches)` 합산.
- `EditRequest`는 그대로 사용(msg_service가 `excel_action`/`pdf_action`을 무시).

## 라우팅 (`file_service.py`)

- `_EXTENSION_MAP`에 `.msg → FileType.MSG` 추가 → `detect_file_type`가 업로드를 허용.
  `detect_file_type`의 안내 문구에 `.msg` 추가. `.md`는 **입력으로 허용하지 않는다**.
- `_service_for(path)`:
  - `.msg` → `msg_service`.
  - `.md` → `msg_service`(산출물 재검증 전용 라우팅). 업로드 검증(`detect_file_type`)에는
    `.md`를 넣지 않으므로 사용자가 `.md`를 업로드할 수는 없다.
- `apply_edit`의 **출력 확장자가 입력과 달라지는 첫 사례**다(`.msg` → `.md`). `EditResult.output_path`는
  msg_service가 만든 `.md` 경로를 담는다. 기존 호출부(단일/배치)는 `output_path`만 사용하므로 호환.

## 배치 (`batch_service.py`)

- `scan_folder`의 `_SUPPORTED_SUFFIXES`에 `.msg` 추가.
- `batch_edit`(별도 출력 폴더 + zip):
  - `.msg`는 내용 매치 시 `msg_service.apply_edit`로 정리된 폴더 경로에 `<stem>_redacted.md` 산출.
  - 이름-only 매치(내용 깨끗 + 파일명/폴더명에 키워드): 기존 로직대로 처리 대상. `.msg`는
    “내용 무수정 복사” 대신 **원본을 그대로 복사하되 정리된 파일명**으로 둔다
    (`.msg` 확장자 유지 — 내용 편집이 아니므로 `.md` 변환하지 않음).
    → 즉 `.msg`의 `.md` 변환은 **내용 매치가 있을 때만** 일어난다.
  - 폴더 세그먼트·파일명 정리·충돌 회피는 기존 `name_redactor`로 동일 적용.
- `batch_edit_in_place`(제자리 교체):
  - **Phase 1(내용)**: `.msg`는 제자리 교체 대상에서 제외한다. 스캔되지만 편집·백업하지 않고,
    결과 항목에 안내(`error`가 아닌 별도 메모)로 “제자리 모드는 .msg 내용을 지원하지 않습니다
    (별도 출력 폴더 모드를 사용하세요)”를 남긴다.
  - **Phase 2(이름)**: `.msg` **파일명에 키워드가 있으면 rename은 그대로 적용**한다(형식이
    바뀌지 않는 순수 rename이므로 안전, 파일명 누출 방지). 폴더 rename도 동일.

## UI (`app.py`)

- `st.file_uploader`의 `type`에 `"msg"` 추가.
- **단일 파일**:
  - `.msg` 검사 결과는 `msg_matches`를 표로 미리보기(필드·줄·키워드·문맥). Excel/PDF 삭제 방식
    선택 UI는 표시하지 않는다(항상 키워드 제거).
  - 승인 → `apply_edit` → `.md` 다운로드. “원본은 .msg, 정리 결과는 .md”임을 안내.
  - 다운로드 파일명 stem은 `name_redactor.redact_filename`으로 정리 + `_redacted.md`.
  - 내용 매치 0건이지만 `.msg` **파일명에 키워드**가 있으면(이름-only) 기존 단일 파일
    이름-only 흐름을 따른다: 원본 `.msg` 바이트를 정리된 파일명으로 다운로드 제공.
- **폴더**: 결과표에 `.msg` 산출물이 `.md`로 표기됨(“변경된 이름” 열에 반영). in-place 모드에서
  `.msg` 내용 미지원 안내 문구 표시.
- PDF의 “빈 공간” 경고처럼, `.msg`는 “서식·이미지·첨부 내용은 보존되지 않는 평문 .md”임을 안내.

## 미지원 입력 처리 & 안전 불변식

- 암호화/RMS 보호 `.msg`, 파싱 실패 → `extract_msg` 예외를 잡아 명확한 안내 메시지로 거부.
  배치에서는 파일별 오류로 격리하고 계속(사유를 `error`에 기록, 삼키지 않음).
- 원본 `.msg` 미수정. 산출 `.md`는 항상 재검증(`msg_service.verify`)을 거친다.
- 경로 처리는 `pathlib`만. 업로드 파일명은 `safe_filename`으로 정규화.
- `extract-msg`는 의존성이 큼(oletools·cryptography·beautifulsoup4·RTFDE 등). `pyproject.toml`
  주 의존성에 추가하고 그 사실을 README/UI에 부담 없이 반영(설치 안내 유지).

## 테스트 (pytest, 파일은 실행 시 생성)

`.msg` 픽스처는 레포에 바이너리로 두지 않고 **실행 시 실제 `.msg`를 생성**한다. 검증 결과
`extract_msg.OleWriter`로 최소 유효 `.msg`를 만들 수 있다(핵심: `PidTagMessageClass=IPM.Note`
= `__substg1.0_001A001F`, 그리고 `__properties_version1.0` 스트림).

**`tests/conftest.py` 또는 헬퍼**
```python
def make_msg(path, *, subject="", body="", sender="", **props): ...
    # OleWriter로 유효 .msg 생성. subject/body/sender는 왕복 확인됨.
    # to/cc는 recipient 스토리지가 필요해 생성 픽스처에서는 비어 있을 수 있음 →
    # to/cc 렌더·제거는 MsgContent 순수 테스트로 커버.
```

**`tests/test_msg_service.py`**
- `render_markdown`: 필드 존재/누락(빈 cc·첨부 없음), 첨부 목록, 본문 없음 표기.
- `search`(MsgContent 경유 또는 생성 픽스처): 제목·본문·보낸사람·첨부 이름의 키워드 검출,
  field/줄번호/문맥 정확성, CONTAINS/EXACT, 대소문자.
- `apply_edit`: `<stem>_redacted.md` 생성, 키워드 제거된 텍스트, 원본 `.msg` 미수정,
  파일명 정리(name_redactor) 반영.
- `verify`: 산출 `.md`에 키워드 잔존 없음(clean=True), 일부러 남긴 경우 clean=False.
- 미지원/손상 `.msg` → 명확한 예외/거부.
- **통합 1개**: 생성한 실제 `.msg`로 path → search → apply_edit → verify 왕복.

**`tests/test_file_service.py` (확장)**
- `.msg` 업로드 허용, `.md`는 입력 거부, `_service_for`가 `.msg`/`.md`를 msg_service로 라우팅.

**`tests/test_batch_service.py` (확장)**
- `scan_folder`가 `.msg`를 포함.
- `batch_edit`: `.msg` 내용 매치 → 정리 폴더에 `<stem>_redacted.md` 산출.
- `batch_edit_in_place`: `.msg` 내용 미처리(원본 유지 + 안내), 파일명 키워드 시 rename 적용.

**`tests/test_app_smoke.py`**: import 무결성 유지.

## 미해결/후속 (범위 밖 기록)

- 첨부파일 내용 검사(엑셀·PDF 첨부를 꺼내 기존 서비스로 처리)는 후속 기능 후보.
- HTML 본문의 고급 정제(표·링크 보존)는 하지 않음. 평문화만.
- `to`/`cc`를 recipient 스토리지까지 포함해 생성하는 픽스처는 필요 시 후속(현재는 MsgContent
  순수 테스트로 커버).
- 서명된(smime) `.msg`의 원본 본문 추출 한계는 안내로 처리.
