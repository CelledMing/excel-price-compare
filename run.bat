@echo off
setlocal
chcp 65001 >nul
pushd "%~dp0"
if not exist "runtime\python.exe" (
    echo Portable runtime is missing. Download the windows-x64 ZIP from GitHub Releases.
    echo For source-code setup, see README.md.
    set "run_status=1"
    goto finish
)
"runtime\python.exe" -B -X utf8 "launcher.py" %*
set "run_status=%errorlevel%"
:finish
echo.
if not "%EXCEL_COMPARE_NO_PAUSE%"=="1" pause
popd
exit /b %run_status%
