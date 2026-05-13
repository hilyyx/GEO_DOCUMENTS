@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Create venv first: python -m venv .venv
  exit /b 1
)

call .venv\Scripts\python.exe -m pip install -q -r requirements.txt -r requirements-build.txt

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

call .venv\Scripts\python.exe -m PyInstaller --noconfirm GEO_Documents.spec

echo.
echo Result: dist\GEO_Documents.exe
endlocal
