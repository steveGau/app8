@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set "PY=%~dp0.venv\Scripts\python.exe"

if exist "%PY%" (
    "%PY%" -c "import sys" >nul 2>&1
    if errorlevel 1 (
        echo This Python environment was copied from another computer. Rebuilding it...
        rmdir /s /q "%~dp0.venv"
    )
)

if not exist "%PY%" (
    echo Creating Python environment...
    where py >nul 2>&1 && py -3.12 -m venv "%~dp0.venv"
    if not exist "%PY%" python -m venv "%~dp0.venv"
    if not exist "%PY%" goto fail
    "%PY%" -m pip install --retries 0 --timeout 20 --upgrade pip
    if errorlevel 1 goto fail
    "%PY%" -m pip install --retries 0 --timeout 20 -r "%~dp0requirements.txt"
    if errorlevel 1 goto fail
)

"%PY%" "%~dp0processExcel.py" %*
if errorlevel 1 goto fail
exit /b 0

:fail
echo processExcel failed.
pause
exit /b 1
