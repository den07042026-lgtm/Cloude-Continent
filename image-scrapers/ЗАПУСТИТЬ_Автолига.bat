@echo off
echo Close "Top-500 VB.xlsx" before running!
echo.
echo 1 - Avtoliga (browser, images from site)
echo 2 - Mikado local (copy from На сортировку folder)
echo 3 - Mikado web (download from mikado-parts.ru)
echo.
set /p choice=Choose (1, 2 or 3):
echo.
if "%choice%"=="1" uv run "%~dp0avtoliga_images.py"
if "%choice%"=="2" uv run --with openpyxl "%~dp0mikado_images.py"
if "%choice%"=="3" uv run "%~dp0mikado_web_images.py"
echo.
pause
