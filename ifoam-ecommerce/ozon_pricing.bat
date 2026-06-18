@echo off
chcp 65001 >nul
set PYTHONUTF8=1
"C:\Users\Admin\.local\bin\uv.exe" run --with openpyxl "C:\Users\Admin\Documents\Ecommerce\ozon_pricing.py" %*
pause
