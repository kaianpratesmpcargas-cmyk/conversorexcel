@echo off
cd /d "%~dp0"
set PY=%LOCALAPPDATA%\Programs\Python\Python314\python.exe
if not exist "%PY%" set PY=python
"%PY%" launcher.py
pause
