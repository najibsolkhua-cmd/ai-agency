@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set PY=py -3) else (set PY=python)
set TO=
set /p TO=Куда отправить тест (Enter — на свою почту): 
%PY% mailer.py --test %TO%
pause
