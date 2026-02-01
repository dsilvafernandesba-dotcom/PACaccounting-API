@echo off
cd /d "%~dp0"
py -m uvicorn api:app --reload
pause
