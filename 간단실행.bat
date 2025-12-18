@echo off
chcp 65001 > nul
title 원가계산서 - 간단 실행

:: py 명령어로 시도 (Microsoft Store Python)
py -m streamlit run app.py 2>nul
if %errorlevel% equ 0 exit /b 0

:: python 명령어로 시도
python -m streamlit run app.py 2>nul
if %errorlevel% equ 0 exit /b 0

:: 둘 다 실패한 경우
echo Python을 찾을 수 없습니다.
echo.
echo 해결 방법:
echo 1. Python을 설치하세요
echo 2. 또는 실행.bat 파일을 사용하세요
echo.
pause




