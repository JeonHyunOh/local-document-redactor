"""폴더(여러 파일) 배치 검사·편집.

로컬 폴더를 스캔해 지원 파일을 찾고, 전체 검사 → 일괄 승인 → 일괄 편집을 수행한다.
한 파일의 실패가 전체 배치를 중단시키지 않도록 파일별로 오류를 격리한다(원본 무수정).
출력은 스캔 루트의 하위 폴더 구조를 그대로 재현한다.
"""

from __future__ import annotations

import shutil
import tempfile
from collections.abc import Callable
from pathlib import Path

from . import file_service, name_redactor
from .models import (
    BatchEditItem,
    BatchSearchItem,
    EditRequest,
    SearchCriteria,
)

# 진행률 콜백: (완료 개수, 전체 개수, 현재 파일 상대경로) -> None
ProgressCallback = Callable[[int, int, str], None]

_SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".pdf"}


def scan_folder(root: Path, recursive: bool = True) -> list[Path]:
    """폴더에서 지원 파일(.xlsx/.xlsm/.pdf)을 찾아 정렬해 반환한다.

    Office 임시 잠금 파일(``~$``)과 숨김 파일은 제외한다. 폴더가 아니면 오류.
    """
    if not root.exists():
        raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {root}")
    if not root.is_dir():
        raise NotADirectoryError(f"폴더가 아닙니다: {root}")

    globber = root.rglob("*") if recursive else root.glob("*")
    files = [
        p
        for p in globber
        if p.is_file()
        and p.suffix.lower() in _SUPPORTED_SUFFIXES
        and not p.name.startswith("~$")
        and not p.name.startswith(".")
    ]
    return sorted(files)


def batch_search(
    root: Path,
    criteria: SearchCriteria,
    recursive: bool = True,
    on_progress: ProgressCallback | None = None,
) -> list[BatchSearchItem]:
    """폴더 내 모든 지원 파일을 검사한다(파일 무수정). 파일별 오류는 격리한다.

    on_progress가 주어지면 파일 하나를 처리할 때마다 (완료, 전체, 상대경로)로 호출한다.
    """
    files = scan_folder(root, recursive)
    total = len(files)
    items: list[BatchSearchItem] = []
    for index, path in enumerate(files, start=1):
        relative = path.relative_to(root).as_posix()
        try:
            report = file_service.search(path, criteria)
            items.append(
                BatchSearchItem(path=str(path), relative_path=relative, report=report)
            )
        except Exception as exc:  # noqa: BLE001 - 사유 기록 후 다음 파일 계속
            items.append(
                BatchSearchItem(path=str(path), relative_path=relative, error=str(exc))
            )
        if on_progress:
            on_progress(index, total, relative)
    return items


def keyword_summary(items: list[BatchSearchItem]) -> dict[str, int]:
    """키워드별 총 발견 건수를 집계한다(입력한 키워드 중 무엇이 얼마나 나왔는지)."""
    counts: dict[str, int] = {}
    for item in items:
        if not item.report:
            continue
        for match in item.report.excel_matches:
            counts[match.keyword] = counts.get(match.keyword, 0) + 1
        for match in item.report.pdf_matches:
            counts[match.keyword] = counts.get(match.keyword, 0) + match.count
    return counts


def match_details(items: list[BatchSearchItem]) -> list[dict[str, str | int]]:
    """어디서 무엇을 찾았는지 행 단위로 펼친다(파일·키워드·위치·내용).

    UI 표시용 평탄화 헬퍼. Excel은 '시트!셀', PDF는 'N페이지'로 위치를 표기한다.
    """
    rows: list[dict[str, str | int]] = []
    for item in items:
        if not item.report:
            continue
        for match in item.report.excel_matches:
            rows.append(
                {
                    "파일": item.relative_path,
                    "키워드": match.keyword,
                    "위치": f"{match.sheet_name}!{match.cell}",
                    "내용": match.original_value,
                }
            )
        for match in item.report.pdf_matches:
            rows.append(
                {
                    "파일": item.relative_path,
                    "키워드": match.keyword,
                    "위치": f"{match.page}페이지",
                    "내용": match.context or f"{match.count}건",
                }
            )
    return rows


def _redact_rel_parent(rel_parent: Path, keywords: list[str], cs: bool) -> Path:
    """상대 부모 경로의 각 세그먼트에서 키워드를 제거한 새 상대 경로를 반환한다."""
    parts = [name_redactor.redact_segment(seg, keywords, cs) for seg in rel_parent.parts]
    return Path(*parts) if parts else Path(".")


def batch_edit(
    root: Path,
    request: EditRequest,
    output_root: Path,
    recursive: bool = True,
    only_with_matches: bool = True,
    on_progress: ProgressCallback | None = None,
) -> list[BatchEditItem]:
    """폴더 내 파일을 일괄 편집해 output_root에 저장한다(폴더·파일명 키워드 정리 포함).

    처리 대상은 '내용 매치 OR 이름 매치'로 판정한다(only_with_matches=True 기준).
    - 내용 매치: 편집 후 정리된 이름으로 저장.
    - 이름만 매치(내용 깨끗): 원본을 정리된 이름으로 복사(내용 무수정).
    출력 경로의 폴더 세그먼트·파일명에서 키워드를 제거하며, 같은 폴더 내 이름 충돌은
    접미사(_1, _2 …)로 회피한다. 각 편집 파일은 재검증하며, 실패 파일은 error만 남긴다.
    on_progress가 주어지면 파일 하나를 처리할 때마다 (완료, 전체, 상대경로)로 호출한다.
    """
    keywords = request.criteria.keywords
    cs = request.criteria.case_sensitive
    files = scan_folder(root, recursive)
    total = len(files)
    items: list[BatchEditItem] = []
    taken_by_dir: dict[Path, set[str]] = {}

    for index, path in enumerate(files, start=1):
        rel = path.relative_to(root)
        relative = rel.as_posix()
        try:
            report = file_service.search(path, request.criteria)
            content_match = report.total_matches > 0
            name_match = name_redactor.name_contains_keyword(path.name, keywords, cs) or any(
                name_redactor.name_contains_keyword(seg, keywords, cs) for seg in rel.parent.parts
            )
            if only_with_matches and not content_match and not name_match:
                items.append(BatchEditItem(path=str(path), relative_path=relative))
                continue

            out_parent = output_root / _redact_rel_parent(rel.parent, keywords, cs)
            out_parent.mkdir(parents=True, exist_ok=True)
            taken = taken_by_dir.setdefault(out_parent, set())

            if content_match:
                edit = file_service.apply_edit(path, request, out_parent)
                verification = file_service.verify(Path(edit.output_path), request.criteria)
                produced = Path(edit.output_path)
                final_name = name_redactor.unique_name(
                    name_redactor.redact_filename(produced.name, keywords, cs), taken
                )
                taken.add(final_name)
                final_path = out_parent / final_name
                if final_path != produced:
                    produced.rename(final_path)
                items.append(
                    BatchEditItem(
                        path=str(path),
                        relative_path=relative,
                        output_path=str(final_path),
                        edit=edit,
                        verification=verification,
                        renamed_to=final_path.relative_to(output_root).as_posix(),
                    )
                )
            else:  # 이름만 매치 → 내용 무수정 복사
                final_name = name_redactor.unique_name(
                    name_redactor.redact_filename(path.name, keywords, cs), taken
                )
                taken.add(final_name)
                final_path = out_parent / final_name
                shutil.copy2(path, final_path)
                items.append(
                    BatchEditItem(
                        path=str(path),
                        relative_path=relative,
                        output_path=str(final_path),
                        renamed_to=final_path.relative_to(output_root).as_posix(),
                    )
                )
        except Exception as exc:  # noqa: BLE001 - 사유 기록 후 다음 파일 계속
            items.append(
                BatchEditItem(path=str(path), relative_path=relative, error=str(exc))
            )
        finally:
            if on_progress:
                on_progress(index, total, relative)
    return items


def batch_edit_in_place(
    root: Path,
    request: EditRequest,
    backup_root: Path,
    recursive: bool = True,
    on_progress: ProgressCallback | None = None,
) -> list[BatchEditItem]:
    """매치가 있는 파일을 편집해 **제자리 교체**하되, 교체 전 원본을 backup_root에 백업한다.

    안전 순서(원본 손실 방지):
    1) 임시 폴더에 편집본을 만든다.
    2) 편집본을 재검증한다. **재검증 실패 시 원본을 건드리지 않고** 유지한다.
    3) 통과한 경우에만 원본을 backup_root에 복사한 뒤 편집본으로 제자리 교체한다.

    backup_root에는 root의 하위 폴더 구조가 그대로 재현된다. output_path는 교체된
    원본 경로(=path)를 가리킨다.
    """
    files = scan_folder(root, recursive)
    total = len(files)
    items: list[BatchEditItem] = []
    for index, path in enumerate(files, start=1):
        relative = path.relative_to(root).as_posix()
        try:
            report = file_service.search(path, request.criteria)
            if report.total_matches == 0:
                items.append(BatchEditItem(path=str(path), relative_path=relative))
                continue

            with tempfile.TemporaryDirectory(prefix="redactor_inplace_") as tmp:
                edit = file_service.apply_edit(path, request, Path(tmp))
                edited_path = Path(edit.output_path)
                verification = file_service.verify(edited_path, request.criteria)

                if not verification.clean:
                    # 재검증 실패 → 원본 유지, 교체하지 않음
                    items.append(
                        BatchEditItem(
                            path=str(path),
                            relative_path=relative,
                            edit=edit,
                            verification=verification,
                            error="재검증 실패로 원본을 그대로 유지했습니다(교체 안 함).",
                        )
                    )
                    continue

                # 원본 백업 후 제자리 교체
                backup_path = backup_root / path.relative_to(root)
                backup_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(path, backup_path)
                shutil.move(str(edited_path), str(path))

            items.append(
                BatchEditItem(
                    path=str(path),
                    relative_path=relative,
                    output_path=str(path),
                    edit=edit,
                    verification=verification,
                )
            )
        except Exception as exc:  # noqa: BLE001 - 사유 기록 후 다음 파일 계속
            items.append(
                BatchEditItem(path=str(path), relative_path=relative, error=str(exc))
            )
        finally:
            if on_progress:
                on_progress(index, total, relative)
    return items
