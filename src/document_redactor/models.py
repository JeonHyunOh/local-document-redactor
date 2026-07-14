"""계층 간 데이터 계약(UI ↔ 서비스 ↔ 향후 LLM).

상태 문자열을 코드 여러 곳에 직접 쓰지 않도록 모든 열거값은 Enum으로 관리한다.
검색과 편집은 별개 단계이므로 요청/결과 모델도 분리한다.
"""

from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field, field_validator


# --------------------------------------------------------------------------- #
# 열거형 (허용된 값의 유일한 출처)
# --------------------------------------------------------------------------- #
class SearchMode(str, Enum):
    """키워드 매칭 방식. 정규식은 MVP UI에 노출하지 않되 추가 가능하도록 열어 둔다."""

    CONTAINS = "contains"
    EXACT = "exact"


class FileType(str, Enum):
    """지원 입력 유형(XLSX/XLSM/PDF/MSG/EML)과 렌더 출력 유형(MD)."""

    XLSX = "xlsx"
    XLSM = "xlsm"
    PDF = "pdf"
    MSG = "msg"
    EML = "eml"
    PPTX = "pptx"
    MD = "md"  # 이메일을 정리해 렌더한 Markdown 산출물(입력 유형 아님)


class ExcelAction(str, Enum):
    """Excel 삭제 방식."""

    REMOVE_KEYWORD = "remove_keyword"  # 셀 값에서 키워드 문자열만 제거
    CLEAR_CELL = "clear_cell"          # 셀 전체 비우기
    DELETE_ROW = "delete_row"          # 키워드가 발견된 행 전체 삭제


class PdfAction(str, Enum):
    """PDF 삭제 방식. MVP는 redaction만 지원한다."""

    REDACT = "redact"


# --------------------------------------------------------------------------- #
# 검색 (파일 무수정 단계)
# --------------------------------------------------------------------------- #
class SearchCriteria(BaseModel):
    """검색 조건. keyword_matcher와 각 서비스가 공유한다."""

    keywords: list[str] = Field(default_factory=list)
    mode: SearchMode = SearchMode.CONTAINS
    case_sensitive: bool = False

    @field_validator("keywords")
    @classmethod
    def _normalize_keywords(cls, value: list[str]) -> list[str]:
        """빈 문자열·공백 전용·중복을 제거하고 입력 순서를 보존한다."""
        seen: set[str] = set()
        cleaned: list[str] = []
        for raw in value:
            keyword = raw.strip()
            if not keyword or keyword in seen:
                continue
            seen.add(keyword)
            cleaned.append(keyword)
        return cleaned


class ExcelMatch(BaseModel):
    """Excel 셀 단위 검색 결과 한 건."""

    file_name: str
    sheet_name: str
    cell: str                 # 예: "B3"
    row: int                  # 행 삭제 처리를 위해 필요
    keyword: str
    original_value: str
    expected_value: str | None = None  # 편집 방식에 따른 예상 변경 값(미리보기용)


class PdfMatch(BaseModel):
    """PDF 검색 결과 한 건."""

    file_name: str
    page: int                 # 1-기반 페이지 번호
    keyword: str
    count: int                # 해당 페이지에서 발견된 개수
    context: str | None = None
    rects: list[tuple[float, float, float, float]] = Field(default_factory=list)  # (x0,y0,x1,y1)


class EmailMatch(BaseModel):
    """이메일을 렌더한 Markdown의 매치 한 건."""

    file_name: str
    field: str                # "제목"/"보낸사람"/"받는사람"/"참조"/"날짜"/"첨부"/"본문"
    line: int                 # 렌더된 .md의 1-기반 줄 번호
    keyword: str
    count: int                # 해당 줄에서 발견된 횟수
    context: str              # 해당 줄 텍스트


class PptxMatch(BaseModel):
    """PowerPoint 슬라이드 텍스트 매치 한 건."""

    file_name: str
    slide: int                # 1-기반 슬라이드 번호
    location: str             # "본문"/"표"/"노트"
    keyword: str
    count: int                # 해당 문단에서 발견된 횟수
    context: str              # 해당 문단 텍스트


class SearchReport(BaseModel):
    """한 파일 검색의 전체 결과. Excel/PDF/이메일/PPTX 중 해당하는 목록만 채운다."""

    file_name: str
    file_type: FileType
    criteria: SearchCriteria
    excel_matches: list[ExcelMatch] = Field(default_factory=list)
    pdf_matches: list[PdfMatch] = Field(default_factory=list)
    email_matches: list[EmailMatch] = Field(default_factory=list)
    pptx_matches: list[PptxMatch] = Field(default_factory=list)
    notes: list[str] = Field(default_factory=list)  # 사용자 안내(예: 텍스트 레이어 없음 가능성)

    @property
    def total_matches(self) -> int:
        return (
            len(self.excel_matches)
            + sum(m.count for m in self.pdf_matches)
            + sum(m.count for m in self.email_matches)
            + sum(m.count for m in self.pptx_matches)
        )


# --------------------------------------------------------------------------- #
# 편집 (사용자 승인 후 단계)
# --------------------------------------------------------------------------- #
class EditRequest(BaseModel):
    """편집 실행 요청. 검색 조건과 파일 유형별 삭제 방식을 함께 담는다."""

    criteria: SearchCriteria
    excel_action: ExcelAction = ExcelAction.REMOVE_KEYWORD
    pdf_action: PdfAction = PdfAction.REDACT


class EditResult(BaseModel):
    """편집 실행 결과. output_path는 원본과 다른 새 파일이어야 한다."""

    source_name: str
    output_path: str
    file_type: FileType
    cells_changed: int = 0
    rows_deleted: int = 0
    redactions_applied: int = 0
    log: list[str] = Field(default_factory=list)


class VerificationResult(BaseModel):
    """편집 후 재검증 결과. 키워드가 남아 있으면 clean=False."""

    output_path: str
    clean: bool
    remaining: SearchReport | None = None


# --------------------------------------------------------------------------- #
# 배치(폴더) 처리 — 여러 파일을 한 번에 검사/편집
# --------------------------------------------------------------------------- #
class BatchSearchItem(BaseModel):
    """폴더 내 파일 한 개의 검색 결과. 처리 실패 시 error에 사유를 담고 계속 진행한다."""

    path: str
    relative_path: str          # 스캔 루트 기준 상대 경로(출력 구조 재현용)
    report: SearchReport | None = None
    error: str | None = None

    @property
    def matches(self) -> int:
        return self.report.total_matches if self.report else 0


class BatchEditItem(BaseModel):
    """폴더 내 파일 한 개의 편집 결과. 실패 파일은 error만 채우고 산출물을 만들지 않는다."""

    path: str
    relative_path: str
    output_path: str | None = None
    edit: EditResult | None = None
    verification: VerificationResult | None = None
    error: str | None = None
    renamed_to: str | None = None  # 이름 정리로 바뀐 최종 경로/파일명(변경 없으면 None)
    note: str | None = None  # 비오류 안내(예: 제자리 모드에서 이메일 내용 미지원)

    @property
    def clean(self) -> bool | None:
        return self.verification.clean if self.verification else None
