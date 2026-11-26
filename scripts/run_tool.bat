@echo off
setlocal
set "ROOT=%~dp0.."
cd /d "%ROOT%"

if not exist ".venv\Scripts\python.exe" (
  echo Environment not found. Run scripts\setup.bat first.
  exit /b 1
)

.venv\Scripts\python.exe src\app.py
