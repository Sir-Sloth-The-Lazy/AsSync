@echo off
REM Daily RISE -> Asana sync. Point Task Scheduler at this file.
REM Edit PROJECT_DIR if you move the folder.

set PROJECT_DIR=%~dp0..
cd /d "%PROJECT_DIR%"

if exist ".venv\Scripts\python.exe" (
    set PY=.venv\Scripts\python.exe
) else (
    set PY=python
)

%PY% run_sync.py >> logs\scheduler.out 2>&1
exit /b %ERRORLEVEL%
