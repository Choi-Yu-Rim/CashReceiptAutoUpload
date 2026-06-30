@echo off
chcp 65001 > nul
echo 현금영수증 자동 업로드 프로그램
echo ================================
echo.
echo 필요한 패키지를 설치하고 있습니다...

pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo.
    echo [오류] pip 설치 실패. Python이 올바르게 설치되어 있는지 확인해주세요.
    pause
    exit /b 1
)

echo.
echo 프로그램을 시작합니다...

python main.py 2>nul
if %errorlevel% neq 0 (
    py main.py
)

pause
