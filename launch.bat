@echo off
setlocal

:: %~dp0 always ends with \  — strip it so quoted paths don't break
set "APP_DIR=%~dp0"
if "%APP_DIR:~-1%"=="\" set "APP_DIR=%APP_DIR:~0,-1%"

set "VENV_PY=%APP_DIR%\venv\Scripts\python.exe"
set "PORT=5000"
set "URL=http://localhost:%PORT%"

:: ------------------------------------------------------------------
:: 1. Locate Python -- prefer the local venv created by deploy.bat
:: ------------------------------------------------------------------
if exist "%VENV_PY%" goto :have_python

echo  No virtual environment found in .\venv\
echo  Please run deploy.bat once to set up the application.
echo.
choice /c YN /m "Run deploy.bat now?"
if %ERRORLEVEL%==1 (
    call "%APP_DIR%\deploy.bat"
    if not exist "%VENV_PY%" (
        echo  deploy.bat did not create the venv. Aborting.
        pause
        exit /b 1
    )
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
:: 3. Start the server in a new console window
::    Note: no /D flag -- we cd inside the cmd call instead, which
::    avoids the trailing-backslash quote-escaping bug in start /D
:: ------------------------------------------------------------------
echo.
echo  Starting SWT CAC Builder...
start "SWT CAC Builder" cmd /c "cd /d "%APP_DIR%" && "%VENV_PY%" app.py & pause"

:: ------------------------------------------------------------------
:: 4. Wait up to 20 seconds for the server to be ready
:: ------------------------------------------------------------------
echo  Waiting for server to be ready...
powershell -NoProfile -Command ^
  "$deadline = (Get-Date).AddSeconds(20);" ^
  "while ((Get-Date) -lt $deadline) {" ^
  "  try { $r = Invoke-WebRequest http://localhost:%PORT%/health -UseBasicParsing -TimeoutSec 1 -EA Stop;" ^
  "    if ($r.StatusCode -eq 200) { exit 0 } } catch {};" ^
  "  Start-Sleep -Milliseconds 400" ^
  "}; exit 1"

if %ERRORLEVEL% neq 0 (
    echo.
    echo  ERROR: Server did not start within 20 seconds.
    echo  Check the "SWT CAC Builder" console window for errors.
    echo.
    pause
    exit /b 1
)

:: ------------------------------------------------------------------
:: 5. Open browser
:: ------------------------------------------------------------------
echo  Server is ready.
start "" "%URL%"
