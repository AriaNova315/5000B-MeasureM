@echo off
setlocal
cd /d "%~dp0"
if errorlevel 1 goto directory_error
if not exist ".venv\Scripts\python.exe" (
  echo Missing .venv. Follow WINDOWS_DEPLOYMENT.md to install offline first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -c "import sys,struct,sysconfig; sys.exit(0 if sys.version_info[:2] == (3,14) and struct.calcsize('P') == 8 and not sysconfig.get_config_var('Py_GIL_DISABLED') else 1)"
if errorlevel 1 (
  echo Expected standard CPython 3.14 x64. Recreate .venv with the documented Python.
  pause
  exit /b 1
)
echo MeasureM: http://127.0.0.1:8002
echo Keep this window open. Press Ctrl+C to stop the server.
".venv\Scripts\python.exe" run.py
set "MEASUREM_EXIT=%ERRORLEVEL%"
if not "%MEASUREM_EXIT%"=="0" (
  echo Startup failed. Check write permission and whether port 8002 is in use.
  pause
)
exit /b %MEASUREM_EXIT%
:directory_error
echo Cannot open the application directory. Use a writable local disk folder.
pause
exit /b 1
