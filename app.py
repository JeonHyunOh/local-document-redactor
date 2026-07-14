"""Streamlit 로컬 UI — 검색 → 미리보기 → 승인 → 편집 → 재검증 흐름.

이 파일에는 비즈니스 로직을 두지 않는다. 모든 파일 검사·편집·검증은
document_redactor 서비스 계층에 위임하고, 여기서는 입력 수집과 결과 표시만 한다.

레이아웃:
- 사이드바: 키워드 · 검색 조건 (모드 공통 입력)
- 본문 탭: [단일 파일] / [폴더 배치]
"""

from __future__ import annotations

import io
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from document_redactor import batch_service, excel_service, file_service, name_redactor
from document_redactor.file_service import UnsupportedFileError
from document_redactor.keyword_matcher import normalize_keywords
from document_redactor.models import (
    EditRequest,
    ExcelAction,
    FileType,
    PdfAction,
    SearchCriteria,
    SearchMode,
)

st.set_page_config(page_title="문서 키워드 검사·삭제 도구", page_icon="🔒", layout="wide")

# 세션 작업 디렉터리(임시). 종료 시 OS가 정리한다.
if "workdir" not in st.session_state:
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="redactor_"))
WORKDIR: Path = st.session_state.workdir
OUTPUT_DIR = WORKDIR / "output"

_EXCEL_ACTION_LABEL = {
    ExcelAction.REMOVE_KEYWORD: "키워드만 제거",
    ExcelAction.CLEAR_CELL: "셀 전체 비우기",
    ExcelAction.DELETE_ROW: "행 전체 삭제",
}
_LABEL_EXCEL_ACTION = {v: k for k, v in _EXCEL_ACTION_LABEL.items()}

# 제자리 모드에서 opt-in으로 완전 삭제할 확장자(내용 정리 불가 형식)
_REMOVAL_SUFFIXES = {".dwg", ".png", ".nwd"}


# --------------------------------------------------------------------------- #
# 사이드바: 키워드 · 검색 조건 (모드 공통)
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("🔑 검색 설정")
    st.text_area(
        "키워드 (한 줄에 하나씩)",
        key="keywords_raw",
        height=140,
        placeholder="대외비\n내부검토용\n주민등록번호",
    )
    keywords = normalize_keywords(st.session_state.get("keywords_raw", "").splitlines())

    st.radio("검색 방식", ["포함", "정확히 일치"], key="mode_label", horizontal=True)
    mode = SearchMode.CONTAINS if st.session_state.get("mode_label") == "포함" else SearchMode.EXACT
    st.checkbox("영문 대소문자 구분", key="case_sensitive")
    case_sensitive = st.session_state.get("case_sensitive", False)

    st.divider()
    if keywords:
        st.caption(f"정규화된 키워드 {len(keywords)}개")
        st.write("· " + "\n· ".join(keywords))
    else:
        st.caption("키워드를 한 줄에 하나씩 입력하세요.")


def build_criteria() -> SearchCriteria:
    return SearchCriteria(keywords=keywords, mode=mode, case_sensitive=case_sensitive)


# --------------------------------------------------------------------------- #
# 헤더
# --------------------------------------------------------------------------- #
st.title("🔒 로컬 엑셀·PDF 키워드 검사 및 삭제 도구")
st.caption(
    "모든 처리는 이 컴퓨터에서만 이루어지며 문서를 외부로 전송하지 않습니다. "
    "원본은 기본적으로 보존되고, 삭제는 검사·미리보기·승인 후에만 실행됩니다."
)

tab_single, tab_folder = st.tabs(["📄 단일 파일", "📁 폴더 (여러 파일)"])


# =========================================================================== #
# 공통 헬퍼
# =========================================================================== #
def _zip_folder(folder: Path) -> bytes:
    """출력 폴더 전체를 zip 바이트로 만든다(구조 유지)."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in sorted(folder.rglob("*")):
            if file.is_file():
                zf.write(file, file.relative_to(folder).as_posix())
    return buffer.getvalue()


# =========================================================================== #
# 단일 파일 모드
# =========================================================================== #
def render_single_file() -> None:
    uploaded = st.file_uploader(
        "파일 업로드 (.xlsx / .xlsm / 텍스트 PDF / .msg / .eml / .pptx)",
        type=["xlsx", "xlsm", "pdf", "msg", "eml", "pptx"],
    )

    if st.button("🔍 검사하기", disabled=not (uploaded and keywords), help="파일을 수정하지 않고 검사만 합니다."):
        for key in ("s_report", "s_saved", "s_edit", "s_verify", "s_editor", "s_approve", "s_all_selected"):
            st.session_state.pop(key, None)
        try:
            saved_path = file_service.save_upload(uploaded.getvalue(), uploaded.name, WORKDIR)
            st.session_state.s_report = file_service.search(saved_path, build_criteria())
            st.session_state.s_saved = saved_path
        except UnsupportedFileError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error("파일을 처리하는 중 오류가 발생했습니다. 파일이 손상되었거나 암호화되었을 수 있습니다.")
            st.exception(exc)

    report = st.session_state.get("s_report")
    if report is None:
        st.info("① 사이드바에 키워드 입력 → ② 파일 업로드 → ③ 검사하기")
        return

    for note in report.notes:
        st.info(note)

    c1, c2 = st.columns(2)
    c1.metric("발견 건수", f"{report.total_matches}건")
    c2.metric("파일 유형", report.file_type.value.upper())

    if report.total_matches == 0:
        up_name = uploaded.name if uploaded else Path(st.session_state.s_saved).name
        if name_redactor.name_contains_keyword(up_name, keywords, case_sensitive):
            clean_name = name_redactor.redact_filename(up_name, keywords, case_sensitive)
            st.info("내용에는 키워드가 없지만 **파일명**에 키워드가 있어 파일명만 정리했습니다.")
            st.caption(f"`{up_name}` → `{clean_name}`")
            st.download_button(
                "📥 파일명 정리본 다운로드 (내용 무수정)",
                data=Path(st.session_state.s_saved).read_bytes(),
                file_name=clean_name,
                use_container_width=True,
            )
        else:
            st.info("발견된 키워드가 없습니다.")
        return

    # 이메일(.msg/.eml): 삭제 방식 선택 없이 승인 → .md 산출
    if report.file_type in (FileType.MSG, FileType.EML):
        st.info("이메일은 서식·이미지·첨부 내용이 보존되지 않는 평문 `.md`로 정리됩니다. "
                "원본 이메일은 수정되지 않습니다.")
        st.dataframe(
            [{"필드": m.field, "줄": m.line, "키워드": m.keyword, "개수": m.count, "문맥": m.context}
             for m in report.email_matches],
            use_container_width=True,
        )
        approved = st.checkbox("위 키워드를 제거한 .md 산출을 승인합니다.", key="s_approve_email")
        if st.button("🗑️ 승인하고 .md 생성", disabled=not approved, type="primary"):
            try:
                request = EditRequest(criteria=report.criteria)
                edit_result = file_service.apply_edit(st.session_state.s_saved, request, OUTPUT_DIR)
                st.session_state.s_edit = edit_result
                st.session_state.s_verify = file_service.verify(Path(edit_result.output_path), report.criteria)
                st.session_state.s_all_selected = True
            except Exception as exc:
                st.error("`.md` 생성 중 오류가 발생했습니다. 결과 파일을 제공하지 않습니다.")
                st.exception(exc)

        edit_result = st.session_state.get("s_edit")
        verification = st.session_state.get("s_verify")
        if edit_result is None or verification is None:
            return

        st.subheader("처리 결과")
        st.metric("제거된 키워드", edit_result.redactions_applied)
        if verification.clean:
            st.success("재검증 완료: 산출 .md에서 키워드가 확인되지 않습니다.")
        else:
            remaining = verification.remaining.total_matches if verification.remaining else 0
            st.error(f"재검증 실패: .md에 키워드가 {remaining}건 남아 있습니다.")

        src_stem = Path(name_redactor.redact_filename(Path(st.session_state.s_saved).name, keywords, case_sensitive)).stem
        dl_name = f"{src_stem}_redacted.md"
        st.download_button(
            "📥 정리된 .md 다운로드",
            data=Path(edit_result.output_path).read_bytes(),
            file_name=dl_name,
            use_container_width=True,
        )
        return

    # PowerPoint(.pptx): 삭제 방식 선택 없이 승인 → 수정본 산출
    if report.file_type is FileType.PPTX:
        st.info("PowerPoint 슬라이드·표·노트의 텍스트에서 키워드를 제거합니다. 원본은 수정되지 않습니다.")
        st.dataframe(
            [{"슬라이드": m.slide, "위치": m.location, "키워드": m.keyword, "개수": m.count, "문맥": m.context}
             for m in report.pptx_matches],
            use_container_width=True,
        )
        approved = st.checkbox("위 키워드 제거를 승인합니다.", key="s_approve_pptx")
        if st.button("🗑️ 승인하고 수정본 생성", disabled=not approved, type="primary"):
            try:
                request = EditRequest(criteria=report.criteria)
                edit_result = file_service.apply_edit(st.session_state.s_saved, request, OUTPUT_DIR)
                st.session_state.s_edit = edit_result
                st.session_state.s_verify = file_service.verify(Path(edit_result.output_path), report.criteria)
                st.session_state.s_all_selected = True
            except Exception as exc:
                st.error("수정본 생성 중 오류가 발생했습니다. 결과 파일을 제공하지 않습니다.")
                st.exception(exc)

        edit_result = st.session_state.get("s_edit")
        verification = st.session_state.get("s_verify")
        if edit_result is None or verification is None:
            return

        st.subheader("처리 결과")
        st.metric("제거된 키워드", edit_result.redactions_applied)
        if verification.clean:
            st.success("재검증 완료: 산출본에서 키워드가 확인되지 않습니다.")
        else:
            remaining = verification.remaining.total_matches if verification.remaining else 0
            st.error(f"재검증 실패: 키워드가 {remaining}건 남아 있습니다.")

        src_stem = Path(name_redactor.redact_filename(Path(st.session_state.s_saved).name, keywords, case_sensitive)).stem
        st.download_button(
            "📥 수정본 다운로드",
            data=Path(edit_result.output_path).read_bytes(),
            file_name=f"{src_stem}_edited.pptx",
            use_container_width=True,
        )
        return

    # 삭제 방식 (Excel만)
    if report.file_type is FileType.PDF:
        st.warning("PDF는 redaction 후 해당 위치에 빈 공간이 남을 수 있습니다.")
        excel_action = None
        base_matches = report.pdf_matches
        rows = [{"삭제": True, "페이지": m.page, "키워드": m.keyword, "개수": m.count, "문맥": m.context or ""} for m in base_matches]
        locked = ["페이지", "키워드", "개수", "문맥"]
    else:
        st.radio("Excel 삭제 방식", list(_LABEL_EXCEL_ACTION), key="excel_label", horizontal=True)
        excel_action = _LABEL_EXCEL_ACTION[st.session_state.get("excel_label", "키워드만 제거")]
        base_matches = report.excel_matches
        previewed = excel_service.preview(report, excel_action)
        rows = [{"삭제": True, "시트": m.sheet_name, "셀": m.cell, "키워드": m.keyword, "원본 값": m.original_value, "예상 값": m.expected_value} for m in previewed]
        locked = ["시트", "셀", "키워드", "원본 값", "예상 값"]

    st.caption("체크된 항목만 삭제됩니다. 기본은 전체 선택이며, 남기고 싶은 항목은 체크를 해제하세요.")
    edited = st.data_editor(pd.DataFrame(rows), disabled=locked, hide_index=True, use_container_width=True, key="s_editor")
    keep_mask = edited["삭제"].tolist()
    selected_matches = [base_matches[i] for i, keep in enumerate(keep_mask) if keep]

    st.divider()
    st.caption(f"선택된 항목: {len(selected_matches)} / {len(base_matches)}")
    can_edit = len(selected_matches) > 0
    approved = st.checkbox("위에서 체크한 항목의 삭제를 승인합니다.", disabled=not can_edit, key="s_approve")
    if st.button("🗑️ 승인하고 삭제본 생성", disabled=not (can_edit and approved), type="primary"):
        try:
            request = EditRequest(criteria=report.criteria, excel_action=excel_action or ExcelAction.REMOVE_KEYWORD, pdf_action=PdfAction.REDACT)
            edit_result = file_service.apply_edit(st.session_state.s_saved, request, OUTPUT_DIR, selected=selected_matches)
            st.session_state.s_edit = edit_result
            st.session_state.s_verify = file_service.verify(Path(edit_result.output_path), report.criteria)
            st.session_state.s_all_selected = len(selected_matches) == len(base_matches)
        except Exception as exc:
            st.error("삭제본 생성 중 오류가 발생했습니다. 결과 파일을 제공하지 않습니다.")
            st.exception(exc)

    edit_result = st.session_state.get("s_edit")
    verification = st.session_state.get("s_verify")
    if edit_result is None or verification is None:
        return

    st.subheader("처리 결과")
    m1, m2, m3 = st.columns(3)
    m1.metric("변경된 셀", edit_result.cells_changed)
    m2.metric("삭제된 행", edit_result.rows_deleted)
    m3.metric("적용된 redaction", edit_result.redactions_applied)

    all_selected = st.session_state.get("s_all_selected", True)
    remaining = verification.remaining.total_matches if verification.remaining else 0
    if verification.clean:
        st.success("재검증 완료: 결과 파일에서 선택한 키워드가 확인되지 않습니다.")
    elif all_selected:
        st.error(f"재검증 실패: 키워드가 {remaining}건 남아 있습니다. 결과 파일을 확인하세요.")
    else:
        st.info(f"선택한 항목을 처리했습니다. 결과 파일에 남은 키워드 발견: {remaining}건 — 선택하지 않은 항목입니다.")

    output_path = Path(edit_result.output_path)
    # 다운로드 파일명도 정리한다: 편집본 접미사(_edited/_redacted)는 유지하고
    # 그 앞의 stem에서만 키워드를 제거한다.
    suffix_tag = "_redacted" if report.file_type is FileType.PDF else "_edited"
    src_name = Path(st.session_state.s_saved).name
    dl_stem = Path(name_redactor.redact_filename(src_name, keywords, case_sensitive)).stem
    dl_final = f"{dl_stem}{suffix_tag}{output_path.suffix}"
    if dl_final != output_path.name:
        st.caption(f"파일명 정리: `{output_path.name}` → `{dl_final}`")
    d1, d2 = st.columns(2)
    d1.download_button("📥 수정본 다운로드", data=output_path.read_bytes(), file_name=dl_final, use_container_width=True)
    d2.download_button("📄 작업 로그 다운로드", data="\n".join(edit_result.log) or "(변경 없음)", file_name=f"{dl_stem}{suffix_tag}_log.txt", use_container_width=True)


# =========================================================================== #
# 폴더(배치) 모드
# =========================================================================== #
def render_folder() -> None:
    c1, c2 = st.columns([4, 1])
    folder_str = c1.text_input("폴더 경로 (로컬)", placeholder=r"예) /Users/이름/문서/검토대상  또는  C:\Users\이름\문서\검토대상")
    recursive = c2.checkbox("하위 폴더 포함", value=True)

    if st.button("🔍 폴더 검사하기", disabled=not (folder_str.strip() and keywords), help="파일을 수정하지 않고 검사만 합니다."):
        for key in ("b_search", "b_edit", "b_root", "b_recursive", "b_in_place", "b_out", "b_backup"):
            st.session_state.pop(key, None)
        root = Path(folder_str).expanduser()
        progress = st.progress(0.0, text="검사 준비 중…")

        def _on_search(done: int, total: int, name: str) -> None:
            progress.progress(done / total if total else 1.0, text=f"검사 중… {done}/{total} — {name}")

        try:
            st.session_state.b_search = batch_service.batch_search(root, build_criteria(), recursive=recursive, on_progress=_on_search)
            st.session_state.b_root = str(root)
            st.session_state.b_recursive = recursive
        except (FileNotFoundError, NotADirectoryError) as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error("폴더를 스캔하는 중 오류가 발생했습니다.")
            st.exception(exc)
        finally:
            progress.empty()

    items = st.session_state.get("b_search")
    if items is None:
        st.info("① 사이드바에 키워드 입력 → ② 폴더 경로 입력 → ③ 폴더 검사하기")
        return

    total_matches = sum(i.matches for i in items)
    errored = [i for i in items if i.error]
    hit_files = [i for i in items if i.matches > 0]

    m1, m2, m3 = st.columns(3)
    m1.metric("스캔한 파일", len(items))
    m2.metric("총 발견 건수", f"{total_matches}건")
    m3.metric("매치된 파일", len(hit_files))

    # 키워드별 발견 건수
    summary = batch_service.keyword_summary(items)
    if keywords:
        st.write("**키워드별 발견 건수**")
        cols = st.columns(len(keywords))
        for col, kw in zip(cols, keywords):
            col.metric(kw, f"{summary.get(kw, 0)}건")

    details = batch_service.match_details(items)
    with st.expander(f"🔎 어디서 무엇을 찾았는지 상세 ({len(details)}건)", expanded=bool(details)):
        st.dataframe(details, use_container_width=True) if details else st.write("발견된 키워드가 없습니다.")

    with st.expander("📁 파일별 요약", expanded=False):
        st.dataframe(
            [{"파일": i.relative_path, "유형": i.report.file_type.value if i.report else "-", "발견 건수": i.matches,
              "상태": "오류: " + i.error if i.error else ("발견" if i.matches else "없음")} for i in items],
            use_container_width=True,
        )
    if errored:
        st.warning(f"{len(errored)}개 파일은 처리할 수 없어 건너뜁니다(암호화·손상·미지원 등). '파일별 요약'을 확인하세요.")

    if total_matches == 0:
        return

    st.divider()
    st.subheader("삭제 실행")
    st.radio("Excel 삭제 방식 (폴더 내 모든 Excel에 적용)", list(_LABEL_EXCEL_ACTION), key="excel_label", horizontal=True)
    excel_action = _LABEL_EXCEL_ACTION[st.session_state.get("excel_label", "키워드만 제거")]
    st.caption("PDF는 항상 redaction으로 처리됩니다. 매치가 있는 파일만 처리합니다.")

    save_mode = st.radio(
        "저장 방식",
        ["별도 출력 폴더 + zip (원본 보존)", "제자리 교체 (원본 백업 후 덮어쓰기)"],
        key="save_mode",
    )
    in_place = save_mode.startswith("제자리")
    root = Path(st.session_state.b_root)
    backup_root = root.parent / f"{root.name}_backup"
    if in_place:
        st.warning(
            f"⚠️ 원본을 덮어씁니다. 교체 전 원본은 다음 위치에 백업됩니다:\n\n`{backup_root}`\n\n"
            "재검증을 통과한 파일만 교체되며, 실패한 파일의 원본은 그대로 유지됩니다. "
            "처음에는 폴더 사본으로 시험해 보시길 권합니다."
        )
        st.checkbox(
            "CAD·이미지·3D 파일(.dwg/.png/.nwd) 완전 삭제 (복구 불가)",
            key="b_remove_targets",
        )
        if st.session_state.get("b_remove_targets"):
            st.warning("⚠️ 체크한 확장자 파일은 **백업 없이 완전 삭제**됩니다. `_removed_log.txt`에만 목록이 남습니다.")

    approve_label = (
        f"위 {len(hit_files)}개 파일의 원본을 덮어쓰는 데 동의합니다(백업 생성됨)."
        if in_place else f"위 {len(hit_files)}개 파일에 대한 일괄 삭제를 승인합니다."
    )
    approved = st.checkbox(approve_label, key="b_approve")
    btn_label = "🗑️ 승인하고 제자리 교체 실행" if in_place else "🗑️ 승인하고 일괄 삭제본 생성"
    if st.button(btn_label, disabled=not approved, type="primary"):
        request = EditRequest(criteria=build_criteria(), excel_action=excel_action, pdf_action=PdfAction.REDACT)
        progress = st.progress(0.0, text="삭제 처리 준비 중…")

        def _on_edit(done: int, total: int, name: str) -> None:
            progress.progress(done / total if total else 1.0, text=f"삭제 처리 중… {done}/{total} — {name}")

        try:
            if in_place:
                st.session_state.b_edit = batch_service.batch_edit_in_place(
                    root, request, backup_root,
                    recursive=st.session_state.b_recursive,
                    on_progress=_on_edit,
                    remove_suffixes=_REMOVAL_SUFFIXES if st.session_state.get("b_remove_targets") else None,
                )
                st.session_state.b_out = None
                st.session_state.b_backup = str(backup_root)
            else:
                out_root = OUTPUT_DIR / "batch"
                st.session_state.b_edit = batch_service.batch_edit(root, request, out_root, recursive=st.session_state.b_recursive, on_progress=_on_edit)
                st.session_state.b_out = str(out_root)
                st.session_state.b_backup = None
            st.session_state.b_in_place = in_place
        except Exception as exc:
            st.error("일괄 처리 중 오류가 발생했습니다.")
            st.exception(exc)
        finally:
            progress.empty()

    edits = st.session_state.get("b_edit")
    if edits is None:
        return

    in_place_done = bool(st.session_state.get("b_in_place"))
    edited = [e for e in edits if e.output_path]
    failed = [e for e in edits if e.error]
    not_clean = [e for e in edited if e.clean is False]

    st.subheader("일괄 처리 결과")
    r1, r2, r3 = st.columns(3)
    r1.metric("제자리 교체" if in_place_done else "수정본 생성", len(edited))
    r2.metric("재검증 실패/유지", len(not_clean) + len(failed))
    r3.metric("매치 없음", len(edits) - len(edited) - len(failed))

    renamed = [e for e in edits if e.renamed_to]
    st.dataframe(
        [{"파일": e.relative_path,
          "변경된 이름": e.renamed_to or "",
          "결과": ("⚠️ " + e.error if e.error
                  else "미생성(매치 없음)" if not e.output_path
                  else ("✅ 교체 완료" if in_place_done else "✅ 검증 통과") if e.clean
                  else "⚠️ 키워드 잔존")} for e in edits],
        use_container_width=True,
    )
    if renamed:
        st.caption(f"이름이 정리된 항목: {len(renamed)}개")
    if in_place_done and renamed:
        st.info(f"이름 변경 기록: `{Path(st.session_state.b_backup) / '_rename_log.txt'}`")

    if not_clean:
        st.error(f"{len(not_clean)}개 파일에서 키워드가 남아 있습니다. 해당 파일을 개별 확인하세요.")
    if failed:
        st.warning(f"{len(failed)}개 파일은 처리하지 못해 원본을 그대로 두었습니다. 위 표의 '⚠️' 사유를 확인하세요.")

    if in_place_done:
        if edited:
            st.success(f"{len(edited)}개 파일을 제자리에서 교체했습니다.")
        removed = [e for e in edits if e.note and "완전 삭제" in e.note]
        if removed:
            st.info(f"완전 삭제된 파일: {len(removed)}개 (복구 불가) — 확장자별 개수: `{Path(st.session_state.b_backup) / '_removed_log.txt'}` (파일명은 기록하지 않음)")
        st.info(f"원본 백업 위치: `{st.session_state.b_backup}`")
    elif edited:
        out_root = Path(st.session_state.b_out)
        st.success(f"수정본이 저장되었습니다: {out_root}")
        st.download_button("📦 결과 폴더 전체 다운로드 (zip)", data=_zip_folder(out_root), file_name="redacted_batch.zip", mime="application/zip")


with tab_single:
    render_single_file()
with tab_folder:
    render_folder()
