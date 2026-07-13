@echo off
REM Windows-Wrapper fuer de_pumpspeicher_v4.py
REM Versucht zuerst die venv, faellt auf System-Python zurueck falls keine venv vorhanden.
REM Aufruf: run_de_pumpspeicher.bat --week  (oder andere Argumente)

SET SCRIPT_DIR=%~dp0
SET VENV=%SCRIPT_DIR%venv\Scripts\python.exe

IF EXIST "%VENV%" (
    "%VENV%" "%SCRIPT_DIR%de_pumpspeicher_v4.py" %*
) ELSE (
    python "%SCRIPT_DIR%de_pumpspeicher_v4.py" %*
)
