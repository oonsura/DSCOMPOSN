@echo off
cd /d "%~dp0"
if exist "%~dp0dscomposn\runx.bat" (
    cd /d "%~dp0dscomposn"
    call "%~dp0dscomposn\runx.bat" %*
) else (
    echo [ERROR] Cannot find dscomposn\runx.bat
    pause
)
