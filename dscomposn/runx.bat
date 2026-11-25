@echo off
cd /d "%~dp0"
title DSCOMPOSN Server

call :findpy
if not defined PY (
    echo [ERROR] Python not found. Install Python 3.8+ first.
    pause
    exit /b 1
)

echo [Python] Using %PY%
%PY% --version

:: Check requirements
%PY% -c "import django, openpyxl, PIL, cv2, waitress, whitenoise" >nul 2>&1
if errorlevel 1 (
    echo [1/4] Installing requirements...
    %PY% -m pip install -q -r requirements.txt
) else (
    echo [1/4] Requirements already satisfied.
)

echo [2/4] Applying migrations...
%PY% manage.py migrate

echo [3/4] Collecting static files...
%PY% manage.py collectstatic --noinput >nul

if /i "%~1"=="--dev" (
    echo [4/4] Starting DEVELOPMENT server at http://127.0.0.1:8000/ (Ctrl+C to stop)
    start "" http://127.0.0.1:8000/
    %PY% manage.py runserver 0.0.0.0:8000
) else (
    if "%DJANGO_DEBUG%"=="" set DJANGO_DEBUG=0
    if "%DJANGO_SERVE_MEDIA%"=="" set DJANGO_SERVE_MEDIA=1
    echo [4/4] Starting PRODUCTION server (Waitress 24 threads) at http://127.0.0.1:8000/ (Ctrl+C to stop)
    echo Login: admin / 1234
    start "" http://127.0.0.1:8000/
    %PY% -m waitress --host=0.0.0.0 --port=8000 --threads=24 dscomposn.wsgi:application
)
pause
exit /b

:findpy
set PY=
if exist "%~dp0venv\Scripts\python.exe" set PY="%~dp0venv\Scripts\python.exe" & exit /b
if exist "%~dp0.venv\Scripts\python.exe" set PY="%~dp0.venv\Scripts\python.exe" & exit /b
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe" & exit /b
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\Python311\python.exe" & exit /b
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\Python310\python.exe" & exit /b
py --version >nul 2>&1 && set PY=py & exit /b
python -c "import sys" >nul 2>&1 && set PY=python & exit /b
for /d %%D in (C:\Python3*) do ( if exist "%%D\python.exe" set PY="%%D\python.exe" & exit /b )
for /d %%D in ("%ProgramFiles%\Python3*") do ( if exist "%%D\python.exe" set PY="%%D\python.exe" & exit /b )
exit /b
