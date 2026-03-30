@echo off
chcp 65001 >nul
:: מוסיף את תיקיית המערכת ל-PATH כדי שffmpeg.exe יימצא
set "PATH=%~dp0;%PATH%"
python "%~dp0server.py"
pause
