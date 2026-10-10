@echo off
rem ============================================================
rem  OCR_Tools launcher (pywebview desktop shell)
rem  ASCII-only comments on purpose: cmd parses this file before
rem  chcp takes effect, non-ASCII comments would execute as junk
rem ============================================================
setlocal EnableExtensions
cd /d "%~dp0"

set "PY="
if exist "..\.venv\Scripts\python.exe" set "PY=..\.venv\Scripts\python.exe"
if not defined PY (
  if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
)
if not defined PY (
  for %%V in (3.13 3.12 3.11 3.10) do (
    if not defined PY (
      py -%%V -c "import sys" >nul 2>&1 && set "PY=py -%%V"
    )
  )
)
if not defined PY (
  where python >nul 2>&1 && set "PY=python"
)
if not defined PY (
  echo [ERROR] Python not found. Install Python 3.10+ first.
  pause
  exit /b 1
)

%PY% -c "import webview, cv2, rapidocr_onnxruntime, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Missing dependencies. Run:
  echo     %PY% -m pip install -r "%~dp0requirements.txt"
  pause
  exit /b 1
)

%PY% "%~dp0run.py"
if errorlevel 1 pause
endlocal
