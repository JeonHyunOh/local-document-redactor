@echo off
REM 로컬 엑셀·PDF 키워드 삭제 도구 실행 스크립트 (Windows)
REM 처음 실행 시 가상환경 생성 + 패키지 설치, 이후에는 바로 실행합니다.
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [오류] Python을 찾을 수 없습니다. https://www.python.org/downloads/ 에서 설치하세요.
    echo        설치 시 "Add Python to PATH"를 반드시 체크하세요.
    pause
    exit /b 1
)

if not exist ".venv" (
    echo [설치] 가상환경을 만들고 패키지를 설치합니다. 잠시 기다려 주세요...
    python -m venv .venv
    call ".venv\Scripts\activate.bat"
    python -m pip install --upgrade pip
    pip install -e .
) else (
    call ".venv\Scripts\activate.bat"
)

echo [실행] 브라우저에서 앱이 열립니다. 종료하려면 이 창에서 Ctrl+C 를 누르세요.
streamlit run app.py
pause
