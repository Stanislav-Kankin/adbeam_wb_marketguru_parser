$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

if (-not (Test-Path '.\.venv\Scripts\python.exe')) {
    throw '.venv not found in project root.'
}

& .\.venv\Scripts\python.exe -m pip install -r requirements-exe.txt
& .\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean wb_inn_gui.spec
