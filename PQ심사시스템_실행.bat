@echo off
cd /d "%~dp0"
start "PQ Server" /min python app.py
timeout /t 3 /nobreak >nul
start http://localhost:5002
pause
