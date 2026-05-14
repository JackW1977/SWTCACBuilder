@echo off
setlocal
set APP_DIR=%~dp0
set VENV_PY=%APP_DIR%venv\Scripts\python.exe
set PORT=5000
set URL=http://localhost:%PORT%

:: ------------------------------------------------------------------
:: 1. Locate Python -- prefer the local venv created by deploy.bat
:: ------------------------------------------------------------------
if exist "%VENV_PY%" (
    set PY=%VENV_PY%
    goto :have_python
)

:: No venv -- offer to run deploy first
echo  No virtual environment found in .\venv\
echo  Please run deploy.bat once to set up the application.
echo.
choice /c YN /m "Run deploy.bat now?"
if %ERRORLEVEL%==1 (
    call "%APP_DIR%deploy.bat"
    if not exist "%VENV_PY%" exit /b 1
    set PY=%VENV_PY%
    goto :have_python
)
exit /b 1

:have_python

:: ------------------------------------------------------------------
:: 2. If already running, just open the browser
:: ------------------------------------------------------------------
powershell -NoProfile -Command ^
  "if (Get-NetTCPConnection -LocalPort %PORT% -EA SilentlyContinue) { exit 0 } else { exit 1 }" >nul 2>&1
if %ERRORLEVEL%==0 (
    echo  SWT CAC Builder is already running at %URL%
    start "" "%URL%"
    exit /b 0
)

:: ------------------------------------------------------------------
:: 3. Start the server in a background window
:: ------------------------------------------------------------------
echo.
echo  Starting SWT CAC Builder...
start "SWT CAC Builder" /D "%APP_DIR%" "%PY%" app.py

:: ------------------------------------------------------------------
:: 4. Wait up to 20 seconds for the server to be ready
:: ------------------------------------------------------------------
echo  Waiting for server to be ready...
powershell -NoProfile -Command ^
  "$deadline = (Get-Date).AddSeconds(20);" ^
  "while ((Get-Date) -lt $deadline) {" ^
  "  if (Get-NetTCPConnection -LocalPort %PORT% -EA SilentlyContinue) { exit 0 };" ^
  "  Start-Sleep -Milliseconds 400" ^
  "};" ^
  "exit 1"

if %ERRORLEVEL% neq 0 (
    echo.
    echo  ERROR: Server did not start within 20 seconds.
    echo  Check the SWT CAC Builder console window for errors.
    echo.
    pause
    exit /b 1
)

:: ------------------------------------------------------------------
:: 5. Open browser
:: ------------------------------------------------------------------
echo  Server is ready.
start "" "%URL%"
