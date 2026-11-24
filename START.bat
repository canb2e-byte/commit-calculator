@echo off
setlocal
cd /d "%~dp0"
echo Commute Collector LOCAL REV 1 / 2026-11-24
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY (
  for /f "delims=" %%P in ('where python 2^>nul ^| findstr /v /i "WindowsApps"') do if not defined PY set "PY="%%P""
)
if not defined PY (
  echo [ERROR] Install Python 3.10 or later from python.org.
  goto :fail
)
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" || goto :fail
%PY% -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
  echo Installing Excel export package...
  %PY% -m pip install -r requirements.txt || goto :fail
)
set "PYTHONUTF8=1"
%PY% commute_tracker.py %* || goto :fail
exit /b 0
:fail
echo [ERROR] Setup or collector failed. Copy this message for troubleshooting.
pause
exit /b 1
