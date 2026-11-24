@echo off
cd /d "%~dp0"
set PY=
py --version >nul 2>&1 && set PY=py
if not defined PY python --version >nul 2>&1 && set PY=python
if not defined PY ( echo Python not found. Install Python 3.8+ first. & pause & exit /b 1 )

%PY% -m pip install -q -r requirements.txt
echo Running automated tests...
%PY% manage.py check
%PY% manage.py test -v 1
echo.
echo Done. (OK = all tests passed)
pause
