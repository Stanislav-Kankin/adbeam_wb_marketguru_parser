@echo off
setlocal
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
  echo [ERROR] .venv not found in project root.
  exit /b 1
)
.\.venv\Scripts\python.exe -m pip install -r requirements-exe.txt
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean wb_inn_gui.spec
endlocal
