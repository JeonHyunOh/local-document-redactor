"""오피스 문서(docx/doc/hwp/hwpx)를 PDF로 변환한다(Word/한글 COM 자동화, pywin32).

COM 의존을 이 모듈에만 둔다. 변환 결과 PDF는 임시 폴더에 캐시하며, 같은 원본(경로·크기·mtime)은
재변환하지 않는다(검사→편집 이중 변환 방지). 변환 뒤 앱을 종료하고, 실패는 ConversionError로
격리한다. Windows + 해당 오피스(Word/한글) 설치가 필요하다.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

_WORD_SUFFIXES = {".docx", ".doc"}
_HWP_SUFFIXES = {".hwp", ".hwpx"}
SUPPORTED_SUFFIXES = _WORD_SUFFIXES | _HWP_SUFFIXES

_cache: dict[tuple[str, int, float], Path] = {}
_cache_dir: Path | None = None


class ConversionError(RuntimeError):
    """오피스 문서를 PDF로 변환하지 못했을 때 발생(환경·보호문서·엔진 오류)."""


def _out_dir() -> Path:
    global _cache_dir
    if _cache_dir is None:
        _cache_dir = Path(tempfile.mkdtemp(prefix="doc2pdf_"))
    return _cache_dir


def convert_to_pdf(path: Path) -> Path:
    """오피스 문서를 PDF로 변환해 그 경로를 반환한다(동일 원본은 캐시 재사용)."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise ConversionError(f"변환 대상이 아닙니다: {suffix or '(확장자 없음)'}")
    stat = path.stat()
    key = (str(path.resolve()), stat.st_size, stat.st_mtime)
    cached = _cache.get(key)
    if cached and cached.exists():
        return cached

    out_pdf = _out_dir() / f"{path.stem}.pdf"
    if suffix in _WORD_SUFFIXES:
        _convert_word(path, out_pdf)
    else:
        _convert_hwp(path, out_pdf)
    if not out_pdf.exists():
        raise ConversionError(f"변환 결과 PDF가 생성되지 않았습니다: {path.name}")
    _cache[key] = out_pdf
    return out_pdf


def _convert_word(src: Path, out_pdf: Path) -> None:
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:  # noqa: BLE001
        raise ConversionError("pywin32/COM을 사용할 수 없습니다(Windows+Word 필요).") from exc

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32.DispatchEx("Word.Application")  # 격리된 새 인스턴스(기존 Word와 충돌 방지)
        for prop, value in (("Visible", False), ("DisplayAlerts", 0)):
            try:
                setattr(word, prop, value)
            except Exception:  # noqa: BLE001 - 비핵심 속성 설정 실패는 무시
                pass
        doc = word.Documents.Open(str(src), ReadOnly=True)
        doc.SaveAs(str(out_pdf), 17)  # wdFormatPDF
        doc.Close(False)
    except Exception as exc:  # noqa: BLE001
        raise ConversionError(f"Word 변환 실패: {src.name} — {exc}") from exc
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:  # noqa: BLE001 - 종료 실패는 무시
                pass
        pythoncom.CoUninitialize()


def _convert_hwp(src: Path, out_pdf: Path) -> None:
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:  # noqa: BLE001
        raise ConversionError("pywin32/COM을 사용할 수 없습니다(Windows+한글 필요).") from exc

    pythoncom.CoInitialize()
    hwp = None
    try:
        hwp = win32.DispatchEx("HwpFrame.HwpObject")
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:  # noqa: BLE001 - 보안모듈 없으면 무시(대화상자 가능)
            pass
        hwp.Open(str(src))
        hwp.SaveAs(str(out_pdf), "PDF", "")
    except Exception as exc:  # noqa: BLE001
        raise ConversionError(f"한글 변환 실패: {src.name} — {exc}") from exc
    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:  # noqa: BLE001
                pass
        pythoncom.CoUninitialize()
