@echo off
REM Windows-Wrapper fuer at_pumpspeicher_v4.py
REM Aktiviert die venv automatisch und startet das Skript.
REM Aufruf: run_at.bat --week  (oder andere Argumente)

SET SCRIPT_DIR=%~dp0
SET VENV=%SCRIPT_DIR%venv\Scripts\activate.bat

IF NOT EXIST "%VENV%" (
    echo Fehler: keine venv gefunden unter %SCRIPT_DIR%venv
    echo Bitte zuerst ausfuehren:
    echo   python -m venv venv
    echo   venv\Scripts\activate
    echo   pip install entsoe-py pandas openpyxl python-dotenv
    exit /b 1
)

CALL "%VENV%"
python "%SCRIPT_DIR%at_pumpspeicher_v4.py" %*
