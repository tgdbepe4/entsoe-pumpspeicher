@echo off
REM Windows-Wrapper fuer at_pumpspeicher_v4.py
REM Versucht zuerst die venv, faellt auf System-Python zurueck falls keine venv vorhanden.
REM Aufruf: run_at_pumpspeicher.bat --week  (oder andere Argumente)

SET SCRIPT_DIR=%~dp0
SET VENV=%SCRIPT_DIR%venv\Scripts\python.exe

IF EXIST "%VENV%" (
    "%VENV%" "%SCRIPT_DIR%at_pumpspeicher_v4.py" %*
) ELSE (
    python "%SCRIPT_DIR%at_pumpspeicher_v4.py" %*
)
