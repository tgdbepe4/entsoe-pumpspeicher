@echo off
REM Windows-Wrapper fuer akw_leistung_ch_v5.py
REM Versucht zuerst die venv, faellt auf System-Python zurueck falls keine venv vorhanden.
REM Aufruf: run_akw.bat --week  (oder andere Argumente)

SET SCRIPT_DIR=%~dp0
SET VENV=%SCRIPT_DIR%venv\Scripts\python.exe

IF EXIST "%VENV%" (
    "%VENV%" "%SCRIPT_DIR%akw_leistung_ch_v5.py" %*
) ELSE (
    python "%SCRIPT_DIR%akw_leistung_ch_v5.py" %*
)
