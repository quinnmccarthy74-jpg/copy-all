@echo off
REM Double-click this to start capturing. Press Ctrl+C in the window to stop.
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    set PY=py
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set PY=python
    ) else (
        echo Python was not found. Install it from https://python.org and tick
        echo "Add python.exe to PATH" during setup, then run this file again.
        pause
        exit /b 1
    )
)

%PY% -c "import win32clipboard, PIL, docx" >nul 2>nul
if not %errorlevel%==0 (
    echo Installing the three required packages, one moment...
    %PY% -m pip install --quiet -r requirements.txt
    if not %errorlevel%==0 (
        echo Install failed. Run this by hand to see why:
        echo    %PY% -m pip install -r requirements.txt
        pause
        exit /b 1
    )
)

%PY% capture_clipboard.py %*
echo.
pause
