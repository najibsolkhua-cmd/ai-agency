@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3"& set "PYW=pyw -3") else (set "PY=python"& set "PYW=pythonw")
%PY% -c "import openpyxl" 2>nul || (echo Устанавливаю нужную библиотеку, подождите... & %PY% -m pip install openpyxl)
start "" %PYW% app.py
