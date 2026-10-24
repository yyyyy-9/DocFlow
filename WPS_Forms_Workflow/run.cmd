@echo off
setlocal
set PYTHONDONTWRITEBYTECODE=1
set PYTHONUTF8=1
if exist "%~dp0environment\python\python.exe" (
  "%~dp0environment\python\python.exe" -X utf8 -B "%~dp0tools\dev_entry.py" %*
) else (
  python -X utf8 -B "%~dp0tools\dev_entry.py" %*
)
exit /b %errorlevel%
