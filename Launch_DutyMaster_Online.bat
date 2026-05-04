@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv .venv
)

echo Installing/updating required packages...
".venv\Scripts\python.exe" -m pip install -r requirements.txt

set DUTYMASTER_PASSWORD=choose-a-password
set DUTYMASTER_SECRET_KEY=make-this-a-long-random-string-change-me
set DUTYMASTER_DATA_DIR=.\data

echo.
echo Starting DutyMaster Online...
echo Open http://127.0.0.1:8000 in your browser.
echo.

".venv\Scripts\python.exe" -m uvicorn main:app --reload

pause
