@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo   Starting Application Build Process...
echo ============================================================

python packaging\build.py
if errorlevel 1 (
    echo [ERROR] Build script exited with an error.
)

pause
