@echo off
setlocal

:: Locate Python 3.11
set PY=C:\Users\jackw\AppData\Local\Programs\Python\Python311\python.exe
if not exist "%PY%" (
    echo Python 3.11 not found at expected path.
    echo Please install Python 3.11 from python.org and update PY in this script.
    pause
    exit /b 1
)

:: Install / upgrade dependencies silently
echo Installing dependencies...
"%PY%" -m pip install -q -r requirements.txt

:: Launch the app
echo.
echo Starting SWT CAC Builder...
"%PY%" app.py

pause
