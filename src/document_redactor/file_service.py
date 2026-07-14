"""파일 유형 판정·안전한 업로드 저장·서비스 라우팅.

UI(app.py)가 excel_service/pdf_service를 직접 알 필요 없이 이 계층을 통해 호출하도록
한다. 업로드 파일명은 그대로 신뢰하지 않고 안전하게 정규화하며, 경로 처리는 pathlib만 쓴다.
"""

from __future__ import annotations

import re
from pathlib import Path

from . import email_service, excel_service, pdf_service, pptx_service
from .models import (
    EditRequest,
    EditResult,
    FileType,
    SearchCriteria,
    SearchReport,
    VerificationResult,
)

_EXTENSION_MAP: dict[str, FileType] = {
    ".xlsx": FileType.XLSX,
    ".xlsm": FileType.XLSM,
    ".pdf": FileType.PDF,
    ".msg": FileType.MSG,
    ".eml": FileType.EML,
    ".pptx": FileType.PPTX,
}
_SAFE_NAME = re.compile(r"[^A-Za-z0-9._가-힣-]+")


class UnsupportedFileError(ValueError):
    """지원하지 않는 파일 형식일 때 발생. 사용자에게 안내 메시지로 표시한다."""


def detect_file_type(filename: str) -> FileType:
    """확장자로 지원 유형을 판정한다. 미지원이면 UnsupportedFileError."""
    suffix = Path(filename).suffix.lower()
    file_type = _EXTENSION_MAP.get(suffix)
    if file_type is None:
        raise UnsupportedFileError(
            f"지원하지 않는 파일 형식입니다: {suffix or '(확장자 없음)'}. "
            "지원 형식은 .xlsx, .xlsm, .pdf, .msg, .eml, .pptx 입니다. "
            "(.xls, 암호화 파일, 스캔 전용 PDF는 지원하지 않습니다.)"
        )
    return file_type


def safe_filename(filename: str) -> str:
    """업로드 파일명에서 경로 구분자·위험 문자를 제거해 안전한 basename을 만든다."""
    base = Path(filename).name  # 디렉터리 성분 제거 (경로 조작 방지)
    cleaned = _SAFE_NAME.sub("_", base).strip("._")
    return cleaned or "upload"


def save_upload(data: bytes, filename: str, dest_dir: Path) -> Path:
    """업로드 바이트를 안전한 이름으로 dest_dir에 저장하고 경로를 반환한다."""
    detect_file_type(filename)  # 저장 전에 형식 검증
    dest_dir.mkdir(parents=True, exist_ok=True)
    path = dest_dir / safe_filename(filename)
    path.write_bytes(data)
    return path


def _service_for(path: Path):
    """파일 경로에 맞는 서비스 모듈을 반환한다.

    산출물 재검증을 위해 .md는 email_service로 라우팅한다(업로드 검증에는 .md를
    포함하지 않으므로 사용자가 .md를 입력할 수는 없다).
    """
    suffix = path.suffix.lower()
    if suffix == ".md":
        return email_service
    file_type = detect_file_type(path.name)
    if file_type in (FileType.MSG, FileType.EML):
        return email_service
    if file_type is FileType.PPTX:
        return pptx_service
    if file_type is FileType.PDF:
        return pdf_service
    return excel_service


def search(path: Path, criteria: SearchCriteria) -> SearchReport:
    """파일 유형에 맞는 서비스로 검색을 위임한다(파일 무수정)."""
    return _service_for(path).search(path, criteria)


def apply_edit(path: Path, request: EditRequest, output_dir: Path, selected=None) -> EditResult:
    """파일 유형에 맞는 서비스로 편집을 위임한다(승인 후 실행).

    selected가 주어지면 선택된 항목만 처리한다(Excel은 ExcelMatch, PDF는 PdfMatch 목록).
    """
    return _service_for(path).apply_edit(path, request, output_dir, selected=selected)


def verify(output_path: Path, criteria: SearchCriteria) -> VerificationResult:
    """파일 유형에 맞는 서비스로 재검증을 위임한다."""
    return _service_for(output_path).verify(output_path, criteria)
