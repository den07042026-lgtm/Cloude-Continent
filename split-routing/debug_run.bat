@echo off
set PYTHON=C:\Users\Admin\AppData\Roaming\uv\python\cpython-3.14.3-windows-x86_64-none\python.exe
echo Split Routing - diagnostic run
echo.
"%PYTHON%" --version
echo.
"%PYTHON%" -c "import tkinter; print('tkinter: OK')"
echo.
echo Running app...
"%PYTHON%" "C:\Users\Admin\Documents\SplitRouting\split_routing.py"
echo.
echo Exit code: %ERRORLEVEL%
echo.
if exist "C:\Users\Admin\Documents\SplitRouting\crash.log" (
    echo --- crash.log ---
    type "C:\Users\Admin\Documents\SplitRouting\crash.log"
)
pause
