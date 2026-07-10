#!/usr/bin/env bash
# 로컬 엑셀·PDF 키워드 삭제 도구 실행 스크립트 (macOS / Linux)
# 처음 실행 시 가상환경 생성 + 패키지 설치, 이후에는 바로 실행합니다.
set -e
cd "$(dirname "$0")"

PY="$(command -v python3 || command -v python || true)"
if [ -z "$PY" ]; then
    echo "[오류] Python을 찾을 수 없습니다. https://www.python.org/downloads/ 에서 설치하세요."
    exit 1
fi

if [ ! -d ".venv" ]; then
    echo "[설치] 가상환경을 만들고 패키지를 설치합니다. 잠시 기다려 주세요..."
    "$PY" -m venv .venv
    source .venv/bin/activate
    python -m pip install --upgrade pip
    pip install -e .
else
    source .venv/bin/activate
fi

echo "[실행] 브라우저에서 앱이 열립니다. 종료하려면 Ctrl+C 를 누르세요."
streamlit run app.py
