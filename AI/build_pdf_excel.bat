@echo off
chcp 65001 > nul
echo ============================================
echo  PDF → Excel 변환기  .exe 빌드 스크립트
echo ============================================
echo.

:: Python 설치 확인
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo https://www.python.org 에서 Python 3.10 이상을 설치하세요.
    pause
    exit /b 1
)

echo [1/3] 필요한 패키지 설치 중...
pip install pdfplumber pymupdf openpyxl pillow pyinstaller --quiet
if errorlevel 1 (
    echo [오류] 패키지 설치 실패
    pause
    exit /b 1
)

echo [2/3] .exe 파일 빌드 중... (2~3분 소요)
cd /d "%~dp0"
pyinstaller --onefile --windowed ^
    --name "PDF_Excel_변환기" ^
    --hidden-import=pdfplumber ^
    --hidden-import=fitz ^
    --hidden-import=openpyxl ^
    --hidden-import=PIL ^
    pdf_to_excel_app.py

if errorlevel 1 (
    echo [오류] 빌드 실패
    pause
    exit /b 1
)

echo.
echo [3/3] 완료!
echo.
echo 실행 파일 위치:
echo    %~dp0dist\PDF_Excel_변환기.exe
echo.
echo 이 파일을 누구에게나 배포할 수 있습니다.
echo Python 설치 없이 바로 실행 가능합니다.
echo.
pause
