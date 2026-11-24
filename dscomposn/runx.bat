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
%PY% -c "import django, openpyxl, PIL, cv2" >nul 2>&1
if errorlevel 1 (
    echo [1/3] Installing requirements...
    %PY% -m pip install -q -r requirements.txt
) else (
    echo [1/3] Requirements already satisfied.
)

echo [2/3] Applying migrations...
%PY% manage.py migrate

echo [3/3] Starting server at http://127.0.0.1:8000/ (Ctrl+C to stop)
echo Login: admin / 1234
start "" http://127.0.0.1:8000/
%PY% manage.py runserver 0.0.0.0:8000
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
