@echo off
cd /d "%~dp0"
call :findpy
if not defined PY ( echo Python not found. Install Python 3.8+ first. & pause & exit /b 1 )

echo [1/3] Installing requirements...
%PY% -m pip install -q -r requirements.txt
echo [2/3] Applying migrations...
%PY% manage.py migrate
echo [3/3] Starting server at http://127.0.0.1:8000/  (Ctrl+C to stop)
echo Login: admin / 1234
start "" http://127.0.0.1:8000/
%PY% manage.py runserver 127.0.0.1:8000
pause
exit /b

:findpy
set PY=
py --version >nul 2>&1 && set PY=py
if not defined PY python --version >nul 2>&1 && set PY=python
exit /b
