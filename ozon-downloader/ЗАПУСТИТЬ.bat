@echo off
powershell -noprofile -ExecutionPolicy Bypass -File "%~dp0start_chrome.ps1"
if errorlevel 1 (
    echo.
    echo Chrome failed to start. See messages above.
    pause
    exit /b 1
)
"%~dp0.venv\Scripts\python.exe" "%~dp0ozon_downloader.py"
pause
