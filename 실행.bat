@echo off
chcp 65001 > nul
title 원가계산서 시스템 실행
color 0A

echo.
echo ============================================================
echo          원가계산서 작성 시스템 실행
echo ============================================================
echo.

:: Python 설치 확인 (여러 방법 시도)
set PYTHON_CMD=
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
    goto :python_found
)

py --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=py
    goto :python_found
)

python3 --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python3
    goto :python_found
)

:: Python을 찾지 못한 경우
echo [오류] Python이 설치되지 않았거나 PATH에 등록되지 않았습니다.
echo.
echo 해결 방법:
echo 1. Python을 설치하세요 (python.org 또는 Microsoft Store)
echo 2. 설치 시 "Add Python to PATH" 옵션을 체크하세요
echo 3. Microsoft Store에서 Python을 설치한 경우, py 명령어를 사용할 수 있습니다
echo.
echo Python 설치 확인 중...
where python >nul 2>&1
if %errorlevel% equ 0 (
    echo [발견] python 명령어를 찾았습니다.
    set PYTHON_CMD=python
    goto :python_found
)
where py >nul 2>&1
if %errorlevel% equ 0 (
    echo [발견] py 명령어를 찾았습니다.
    set PYTHON_CMD=py
    goto :python_found
)
echo.
echo Python을 찾을 수 없습니다. Python을 설치해주세요.
pause
exit /b 1

:python_found
echo [확인] Python 설치 확인 완료
%PYTHON_CMD% --version
echo.

:: 필요한 라이브러리 설치
echo [설치] 필요한 라이브러리를 확인하고 설치합니다...
%PYTHON_CMD% -m pip install -q streamlit pandas openpyxl
if %errorlevel% neq 0 (
    echo [경고] 라이브러리 설치 중 문제가 발생했습니다.
    echo 수동으로 설치를 시도합니다...
    %PYTHON_CMD% -m pip install streamlit pandas openpyxl
    if %errorlevel% neq 0 (
        echo [오류] 라이브러리 설치에 실패했습니다.
        echo 수동으로 다음 명령어를 실행해주세요:
        echo %PYTHON_CMD% -m pip install streamlit pandas openpyxl
        pause
        exit /b 1
    )
)

echo.
echo ============================================================
echo          프로그램을 실행합니다...
echo          브라우저가 자동으로 열립니다.
echo ============================================================
echo.
echo 종료하려면 이 창에서 Ctrl+C를 누르세요.
echo.

:: Streamlit 앱 실행
%PYTHON_CMD% -m streamlit run app.py --server.headless false

:: 에러 발생 시
if %errorlevel% neq 0 (
    echo.
    echo ============================================================
    echo [오류] 프로그램 실행 중 문제가 발생했습니다.
    echo 위의 오류 메시지를 확인해주세요.
    echo ============================================================
    echo.
    pause
)

exit /b 0
