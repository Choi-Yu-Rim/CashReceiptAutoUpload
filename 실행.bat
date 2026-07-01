@echo off
echo Cash Receipt Auto Upload
echo ========================
echo.
echo Installing packages...

python -m pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    py -m pip install -r requirements.txt --quiet
    if %errorlevel% neq 0 (
        echo.
        echo [ERROR] pip install failed. Please check Python installation.
        pause
        exit /b 1
    )
)

echo.
echo Starting program...

python main.py
if %errorlevel% neq 0 (
    py main.py
)

pause
