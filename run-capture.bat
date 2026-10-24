@echo off
REM Double-click this to start capturing. Press Ctrl+C in the window to stop.
setlocal
cd /d "%~dp0"

REM --- find a Python launcher -------------------------------------------
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY goto :no_python

REM --- install the dependencies if they are not importable yet -----------
REM Note: use "if errorlevel N", which is evaluated when the line RUNS.
REM "%errorlevel%" would be substituted at parse time and read a stale value.
%PY% -c "import win32clipboard, PIL, docx" >nul 2>nul
if not errorlevel 1 goto :run

echo Installing the three required packages, one moment...
%PY% -m pip install --quiet --no-warn-script-location -r requirements.txt
%PY% -c "import win32clipboard, PIL, docx" >nul 2>nul
if errorlevel 1 goto :install_failed
echo Done.
echo.

REM --- go ----------------------------------------------------------------
:run
%PY% capture_clipboard.py %*
echo.
pause
exit /b 0

:no_python
echo Python was not found.
echo.
echo Install it from https://python.org (or the Microsoft Store), then open a
echo NEW terminal window and run this file again.
echo.
pause
exit /b 1

:install_failed
echo.
echo The packages could not be installed. Run this by hand to see the reason:
echo.
echo     %PY% -m pip install -r requirements.txt
echo.
echo A "not on PATH" warning about pywin32_postinstall.exe is harmless and is
echo not the problem.
echo.
pause
exit /b 1
