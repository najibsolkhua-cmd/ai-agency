@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set PY=py -3) else (set PY=python)
echo Устанавливаю библиотеку openpyxl...
%PY% -m pip install --upgrade openpyxl
echo.
echo Готово. Если выше нет красных ошибок — можно переходить к следующему шагу.
pause
