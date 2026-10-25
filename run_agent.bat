@echo off
title YouTube Websites Extractor Agent
cd /d "%~dp0"

echo ============================================================
echo   YouTube Websites Extractor & Categorizer Agent
echo ============================================================
echo.
echo [1] Press Enter to process Default Channel (@setupsaitony)
echo [2] Or paste any YouTube Channel URL or Handle
echo.
set /p TARGET_CHANNEL="Enter Channel URL or Handle [Default: @setupsaitony]: "

if "%TARGET_CHANNEL%"=="" set TARGET_CHANNEL=@setupsaitony

echo.
echo [*] Launching Python Extraction & Categorization Pipeline...
python extract_and_categorize_websites.py "%TARGET_CHANNEL%"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Script encountered an error (Code: %ERRORLEVEL%)
) else (
    echo.
    echo ============================================================
    echo   Extraction & Categorization Completed Successfully!
    echo ============================================================
)

echo.
echo Press any key to exit...
pause >nul
