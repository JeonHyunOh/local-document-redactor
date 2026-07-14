"""폴더(여러 파일) 배치 검사·편집.

로컬 폴더를 스캔해 지원 파일을 찾고, 전체 검사 → 일괄 승인 → 일괄 편집을 수행한다.
한 파일의 실패가 전체 배치를 중단시키지 않도록 파일별로 오류를 격리한다(원본 무수정).
출력은 스캔 루트의 하위 폴더 구조를 그대로 재현한다.
"""

from __future__ import annotations

import shutil
import tempfile
from collections import Counter
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

_SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".pdf", ".msg", ".eml", ".pptx"}
_EMAIL_SUFFIXES = {".msg", ".eml"}
_REMOVED_LOG_NAME = "_removed_log.txt"  # 대상 폴더에 남기는 삭제 기록(도구 산출물, rename 대상 제외)


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


def scan_all_files(root: Path, recursive: bool = True) -> list[Path]:
    """숨김·Office 잠금(~$) 파일을 제외한 모든 파일(확장자 무관)을 정렬해 반환한다.

    확장자 무관 파일명 정리·형식 제거가 지원 형식에 국한되지 않도록 하는 스캔이다.
    """
    globber = root.rglob("*") if recursive else root.glob("*")
    return sorted(
        p
        for p in globber
        if p.is_file() and not p.name.startswith("~$") and not p.name.startswith(".")
    )


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


def _disk_unique(target_dir: Path, name: str) -> str:
    """target_dir 안에서 충돌하지 않는 이름을 만든다(현재 디렉터리 항목 기준)."""
    taken = {p.name for p in target_dir.iterdir()} if target_dir.exists() else set()
    return name_redactor.unique_name(name, taken)


def batch_edit_in_place(
    root: Path,
    request: EditRequest,
    backup_root: Path,
    recursive: bool = True,
    on_progress: ProgressCallback | None = None,
    remove_suffixes: set[str] | None = None,
) -> list[BatchEditItem]:
    """매치가 있는 파일을 편집해 제자리 교체하고(백업 후), 이름의 키워드도 제거한다.

    Phase 0(opt-in): remove_suffixes가 주어지면 해당 확장자 파일을 **완전 삭제**한다
             (백업 없음, backup_root/_removed_log.txt에 목록 기록, 복구 불가). 기본은 미실행.
    Phase 1: 내용 편집 → 재검증 통과 시에만 백업 후 제자리 교체(기존 안전 순서 유지).
    Phase 2: 디스크에서 파일→폴더(bottom-up) 순으로 이름의 키워드를 제거해 rename한다.
             파일 rename은 **확장자 무관 모든 파일** 대상. 루트 폴더 자체는 제외.
             rename 전 원본 백업을 보장하고, 모든 rename을 backup_root/_rename_log.txt에 남긴다.

    backup_root에는 root의 하위 폴더 구조가 그대로 재현된다. 내용 편집 파일의
    output_path는 교체된 원본 경로를, 이름이 바뀐 파일은 최종 경로를 가리킨다.
    """
    keywords = request.criteria.keywords
    cs = request.criteria.case_sensitive
    files = scan_folder(root, recursive)
    total = len(files)
    items: list[BatchEditItem] = []
    by_path: dict[str, BatchEditItem] = {}

    # --- Phase 0: 형식 제거(opt-in, 완전 삭제·백업 없음) ---
    # 로그에는 확장자별 삭제 개수만 남기고 파일명은 기록하지 않는다(삭제 대상 파일명 자체가
    # 민감정보일 수 있으므로).
    removed_counts: Counter[str] = Counter()
    if remove_suffixes:
        targets = {s.lower() for s in remove_suffixes}
        for path in scan_all_files(root, recursive):
            suffix = path.suffix.lower()
            if suffix not in targets:
                continue
            rel = path.relative_to(root).as_posix()
            try:
                path.unlink()
                removed_counts[suffix] += 1
                items.append(
                    BatchEditItem(
                        path=str(path),
                        relative_path=rel,
                        note="완전 삭제됨(제거 대상 확장자)",
                    )
                )
            except Exception as exc:  # noqa: BLE001 - 사유 기록 후 계속
                items.append(BatchEditItem(path=str(path), relative_path=rel, error=str(exc)))
    if removed_counts:
        # 삭제 기록은 대상 폴더 안에 남긴다(확장자별 개수만, 파일명 미기록).
        lines = [f"{ext} {count}개 삭제" for ext, count in sorted(removed_counts.items())]
        (root / "_removed_log.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")

    # --- Phase 1: 내용 편집(기존 동작) ---
    for index, path in enumerate(files, start=1):
        relative = path.relative_to(root).as_posix()
        try:
            if path.suffix.lower() in _EMAIL_SUFFIXES:
                # 이메일: 키워드가 있으면 정리된 .md를 만들고 원본 이메일을 삭제한다.
                # 원본은 backup_root에 백업하고, .md 재검증 통과 시에만 삭제한다.
                report = file_service.search(path, request.criteria)
                name_hit = name_redactor.name_contains_keyword(path.name, keywords, cs)
                if report.total_matches == 0 and not name_hit:
                    item = BatchEditItem(path=str(path), relative_path=relative)  # 깨끗 → 유지
                    items.append(item)
                    by_path[str(path)] = item
                    continue
                with tempfile.TemporaryDirectory(prefix="redactor_eml_") as tmp:
                    edit = file_service.apply_edit(path, request, Path(tmp))
                    md_tmp = Path(edit.output_path)
                    verification = file_service.verify(md_tmp, request.criteria)
                    if not verification.clean:
                        item = BatchEditItem(
                            path=str(path),
                            relative_path=relative,
                            edit=edit,
                            verification=verification,
                            error="이메일 정리본 재검증 실패로 원본을 그대로 유지했습니다(.md 미생성).",
                        )
                        items.append(item)
                        by_path[str(path)] = item
                        continue
                    backup_path = backup_root / path.relative_to(root)
                    backup_path.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(path, backup_path)
                    clean_stem = Path(name_redactor.redact_filename(path.name, keywords, cs)).stem
                    md_name = _disk_unique(path.parent, f"{clean_stem}_redacted.md")
                    md_final = path.parent / md_name
                    shutil.move(str(md_tmp), str(md_final))
                path.unlink()  # 원본 이메일 삭제(백업은 유지)
                item = BatchEditItem(
                    path=str(path),
                    relative_path=relative,
                    output_path=str(md_final),
                    edit=edit,
                    verification=verification,
                    renamed_to=md_final.relative_to(root).as_posix(),
                    note="이메일 정리본 .md 생성, 원본 이메일 삭제",
                )
                items.append(item)
                by_path[str(path)] = item
                continue
            report = file_service.search(path, request.criteria)
            if report.total_matches == 0:
                item = BatchEditItem(path=str(path), relative_path=relative)
                items.append(item)
                by_path[str(path)] = item
                continue

            with tempfile.TemporaryDirectory(prefix="redactor_inplace_") as tmp:
                edit = file_service.apply_edit(path, request, Path(tmp))
                edited_path = Path(edit.output_path)
                verification = file_service.verify(edited_path, request.criteria)

                if not verification.clean:
                    # 재검증 실패 → 원본 유지, 교체하지 않음
                    item = BatchEditItem(
                        path=str(path),
                        relative_path=relative,
                        edit=edit,
                        verification=verification,
                        error="재검증 실패로 원본을 그대로 유지했습니다(교체 안 함).",
                    )
                    items.append(item)
                    by_path[str(path)] = item
                    continue

                # 원본 백업 후 제자리 교체
                backup_path = backup_root / path.relative_to(root)
                backup_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(path, backup_path)
                shutil.move(str(edited_path), str(path))

            item = BatchEditItem(
                path=str(path),
                relative_path=relative,
                output_path=str(path),
                edit=edit,
                verification=verification,
            )
            items.append(item)
            by_path[str(path)] = item
        except Exception as exc:  # noqa: BLE001 - 사유 기록 후 다음 파일 계속
            item = BatchEditItem(path=str(path), relative_path=relative, error=str(exc))
            items.append(item)
            by_path[str(path)] = item
        finally:
            if on_progress:
                on_progress(index, total, relative)

    # --- Phase 2: 이름 rename (파일 먼저, 그다음 폴더 bottom-up) ---
    renames: list[tuple[str, str]] = []  # (이전상대, 새상대)

    # 2a) 파일 rename — 확장자 무관 모든 파일. Phase 1 실패(error) 파일은 건드리지 않는다.
    for path in scan_all_files(root, recursive):
        if path.name == _REMOVED_LOG_NAME:
            continue  # 도구가 만든 삭제 기록은 이름 정리 대상에서 제외(훼손 방지)
        item = by_path.get(str(path))
        if item is not None and item.error:
            continue
        if not name_redactor.name_contains_keyword(path.name, keywords, cs):
            continue
        cleaned = name_redactor.redact_filename(path.name, keywords, cs)
        if cleaned == path.name:
            continue  # 키워드가 확장자에만 있는 등, 정리해도 이름이 그대로면 rename하지 않음
        # 백업 보장(Phase 1에서 백업 안 된 경우)
        backup_path = backup_root / path.relative_to(root)
        if not backup_path.exists():
            backup_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, backup_path)
        new_name = _disk_unique(path.parent, cleaned)
        new_path = path.parent / new_name
        old_rel = path.relative_to(root).as_posix()
        path.rename(new_path)
        renames.append((old_rel, new_path.relative_to(root).as_posix()))
        if item is None:  # 미지원 확장자 파일 — 결과 항목 신규 생성
            item = BatchEditItem(path=str(path), relative_path=old_rel)
            items.append(item)
        by_path[str(new_path)] = item
        item.output_path = str(new_path)
        item.renamed_to = new_name

    # 2b) 폴더 rename — 깊은 곳부터, 루트 제외
    if recursive:
        dirs = sorted(
            (p for p in root.rglob("*") if p.is_dir()),
            key=lambda p: len(p.relative_to(root).parts),
            reverse=True,
        )
        for d in dirs:
            if not name_redactor.name_contains_keyword(d.name, keywords, cs):
                continue
            new_name = _disk_unique(d.parent, name_redactor.redact_segment(d.name, keywords, cs))
            new_path = d.parent / new_name
            old_rel = d.relative_to(root).as_posix()
            d.rename(new_path)
            renames.append((old_rel + "/", new_path.relative_to(root).as_posix() + "/"))

    # rename 로그 기록(변경이 있을 때만 생성)
    if renames:
        backup_root.mkdir(parents=True, exist_ok=True)
        log_lines = [f"{old} -> {new}" for old, new in renames]
        (backup_root / "_rename_log.txt").write_text("\n".join(log_lines) + "\n", encoding="utf-8")

    return items
