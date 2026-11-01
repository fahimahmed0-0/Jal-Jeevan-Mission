@echo off
setlocal
cd /d "%~dp0"
title JJM PHED Assam - TPIA Abstract of Bill

echo ============================================================
echo   JAL JEEVAN MISSION (JJM) - PHED ASSAM
echo   TPIA Abstract of Bill (Payment Recommendation)
echo ============================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python 3.10+ and try again.
    pause
    exit /b 1
)

REM Install / update dependencies when any of them is missing
python -c "import streamlit, pdfplumber, openpyxl, pypdfium2, rapidocr, onnxruntime, pandas" >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Installing required dependencies...
    python -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to install dependencies.
        pause
        exit /b 1
    )
)

echo [INFO] Starting the app on http://localhost:8501
echo.
python -m streamlit run app.py --server.port 8501

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] The application closed with an error.
    pause
)
