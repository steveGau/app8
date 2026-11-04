@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Creating Python environment...
    python -m venv .venv
    if errorlevel 1 goto fail
    ".venv\Scripts\python.exe" -m pip install -r "%~dp0requirements.txt"
    if errorlevel 1 goto fail
)
".venv\Scripts\python.exe" "%~dp0processExcel.py" %*
if errorlevel 1 goto fail
exit /b 0

:fail
echo processExcel failed.
pause
exit /b 1
