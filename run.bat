@echo off
REM SendToTTS launcher (windowless)
REM v1.1.3

REM Launch pythonw in a new process so this cmd window can close immediately
start "" pythonw "%~dp0main.py"
exit 