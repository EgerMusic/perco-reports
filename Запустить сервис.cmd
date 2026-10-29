@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
  py -3 "%~dp0server.py"
  goto finished
)
where python >nul 2>nul
if not errorlevel 1 (
  python "%~dp0server.py"
  goto finished
)
set "BUNDLED_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%BUNDLED_PYTHON%" (
  "%BUNDLED_PYTHON%" "%~dp0server.py"
  goto finished
)
echo Python 3 was not found. Install Python or ask for a standalone launcher.
:finished
echo.
echo Press any key to close this window.
pause >nul
