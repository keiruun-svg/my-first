@echo off
chcp 65001 > nul
echo ========================================
echo   AJW 연간 발주 계획 생성 시작
echo ========================================
echo.

echo [STEP 1] 3개년 생산자재 사용내역 생성 중...
python "%~dp0STEP1_사용내역_생성.py"
if errorlevel 1 goto error

echo.
echo [STEP 2] 2026 연간 발주계획 생성 중...
python "%~dp0STEP2_발주계획_생성.py"
if errorlevel 1 goto error

echo.
echo ✅ 전체 완료!
goto end

:error
echo ❌ 오류가 발생했습니다. 위 메시지를 확인하세요.

:end
pause
