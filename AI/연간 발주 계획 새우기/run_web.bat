@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo.
echo ========================================
echo   AJW 발주계획 웹앱 시작 중...
echo ========================================
echo.

REM 가상환경 streamlit 우선 사용, 없으면 PATH에서 찾기
set STREAMLIT=.venv\Scripts\streamlit.exe
if not exist "%STREAMLIT%" set STREAMLIT=streamlit

echo 브라우저에서 자동으로 열립니다.
echo 사내 다른 PC에서 접속: http://%COMPUTERNAME%:8501
echo.
echo 종료하려면 이 창에서 Ctrl+C 를 누르세요.
echo.

%STREAMLIT% run web_app.py --server.address=0.0.0.0 --server.port=8501 --server.headless=false --browser.gatherUsageStats=false

pause
