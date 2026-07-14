# 오피스 문서(docx/doc/hwp/hwpx) → PDF 변환 후 검사·삭제 — 설계

- 작성일: 2026-07-14
- 상태: 승인됨 (구현 대기)
- 대상 레포: `local-document-redactor`

## 배경 / 목표

`.docx`/`.doc`(MS Word), `.hwp`/`.hwpx`(한컴 한글) 문서를 **PDF로 변환**한 뒤 기존 `pdf_service`로
키워드·패턴을 검사·redaction한다. 순수 파이썬으로는 이들 형식의 안정적 변환이 불가하므로, 이 PC에
설치된 **Word/한글을 COM 자동화(pywin32)** 로 구동해 PDF를 생성한다.

**검증 완료:** 파이썬 `win32com`(late binding)으로 docx·hwp → PDF 변환이 성공하며, 결과 PDF는
텍스트 검색이 가능(키워드·패턴 redaction 가능)함을 실측했다.

## 핵심 결정 (사용자 승인)

- **결과물**: `<원본이름>_redacted.pdf` (키워드+패턴 제거된 PDF).
- **원본 처리**: 단일·출력본 모드는 보존. **제자리 모드는 원본을 backup 후 삭제**(이메일→md와 동일).
- **형식**: `.docx`, `.doc`(Word), `.hwp`, `.hwpx`(한글) 모두.
- **모드**: 단일 파일 + 폴더(출력본·제자리) 모두.
- **의존성**: `pywin32`. **Windows + Word/한글 설치 필수.** 미충족 시 명확히 거부.

## 비목표
- 편집 가능한 원본 형식으로의 복원(결과는 PDF 고정).
- LibreOffice 등 다른 변환 엔진(미설치). macOS/Linux 지원.
- 비밀번호로 보호된 문서 자동 해제.

## 아키텍처

### 신규 모듈 `doc_converter.py` (COM 격리)
```
convert_to_pdf(path: Path) -> Path
    # 확장자로 변환기 선택. 변환된 PDF의 경로(임시 캐시)를 반환.
    #   .docx/.doc  → Word COM (wdFormatPDF=17)
    #   .hwp/.hwpx  → 한글 COM (SaveAs "PDF", 보안모듈 등록)
    # 캐시: (절대경로, 크기, mtime) 동일하면 재변환하지 않음(search→edit 이중 변환 방지).
    # 변환 후 앱 종료 + 잔여 프로세스 정리. 실패 시 ConversionError.
```
- COM은 이 모듈에만. `win32com.client.Dispatch` late binding. 각 변환은 `CoInitialize`/`CoUninitialize`.
- **프로세스 정리**: 변환 뒤 `Quit()` + 안전장치로 잔여 WINWORD/Hwp 프로세스 종료(모듈이 띄운 것만).
- 미지원 환경(pywin32 없음, Windows 아님, Office 미설치)·보호 문서·변환 실패 → `ConversionError`.

### 신규 서비스 `office_service.py` (search/apply_edit/verify 계약)
```
search(path, criteria) -> SearchReport
    # convert_to_pdf(path) → pdf_service.search(pdf) 결과를 원본 파일명으로 반환 + 변환 안내 note.
apply_edit(path, request, output_dir, selected=None) -> EditResult
    # convert_to_pdf(path) → pdf_service.apply_edit(pdf, ...) → <원본stem>_redacted.pdf로 산출.
verify(output_path, criteria) -> VerificationResult
    # 산출물은 .pdf이므로 pdf_service.verify로 위임.
```
- `doc_converter.convert_to_pdf`는 monkeypatch 가능하게 두어, office_service 로직은 **실제 Office 없이**
  (가짜 PDF 반환) 단위 테스트한다.

### 데이터 모델 (`models.py`)
- `FileType`에 `DOCX`, `DOC`, `HWP`, `HWPX` 추가.

### 라우팅 (`file_service.py`)
- `_EXTENSION_MAP`에 네 형식 추가 → 업로드 허용. `detect_file_type` 안내 문구 갱신.
- `_service_for`: `.docx/.doc/.hwp/.hwpx` → `office_service`. (산출 `.pdf`는 기존대로 pdf_service로 재검증.)

### 배치 (`batch_service.py`)
- `_SUPPORTED_SUFFIXES`에 네 형식 추가.
- `batch_edit`(출력본): 변환→redaction된 `<stem>_redacted.pdf`를 출력 폴더에 산출(원본 보존).
- `batch_edit_in_place`: **이메일과 동일 패턴** — 키워드/패턴이 있으면 `<stem>_redacted.pdf`를 같은
  폴더에 만들고 **원본 문서는 backup 후 삭제**(변환 실패는 격리·원본 유지).

### UI (`app.py`)
- 업로더 `type`에 `docx/doc/hwp/hwpx` 추가. 단일 파일: 변환→검사 미리보기(pdf 매치)→승인→`.pdf` 다운로드.
- **변환은 느릴 수 있음**을 안내(파일마다 Office 구동). 진행률 표시.
- Windows/Office 미설치 등 변환 실패 시 명확한 오류 메시지.

## 성능·안전
- 변환은 파일당 수 초. 대량 폴더는 오래 걸림 → 진행률·안내. (인스턴스 재사용 최적화는 후속.)
- 변환 산출 PDF는 임시 폴더. 작업 후 정리. 원본 삭제는 제자리 모드 + backup 시에만.
- 재검증은 산출 PDF에 대해 키워드+패턴 모두 확인(기존 pdf_service).
- COM 실패·보호 문서·미지원 환경은 파일별 오류로 격리, 배치 계속.

## 테스트
- `tests/test_office_service.py`: `doc_converter.convert_to_pdf`를 monkeypatch(가짜 PDF를 fitz로 생성)
  → search/apply_edit/verify가 pdf_service에 위임하고 `<원본stem>_redacted.pdf`를 만드는지 검증.
- `tests/test_doc_converter.py`: 실제 변환은 `pytest.importorskip("win32com")` + Office COM 가용 시에만
  실행(이 PC에서만). 미가용이면 skip.
- `file_service`/`batch_service` 라우팅·배치 테스트 확장(변환기는 monkeypatch).
- `test_app_smoke.py` 유지.

## 미해결/후속
- 변환 인스턴스 재사용(배치 성능), LibreOffice 대체 엔진, 비밀번호 문서, macOS/Linux.
