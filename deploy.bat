@echo off
setlocal enabledelayedexpansion
set APP_DIR=%~dp0
set VENV_DIR=%APP_DIR%venv
set VENV_PY=%VENV_DIR%\Scripts\python.exe

echo.
echo =====================================================
echo   SWT CAC Builder  --  Local Deployment Setup
echo =====================================================
echo.

:: ------------------------------------------------------------------
:: 1. Locate Python 3.10+
:: ------------------------------------------------------------------
set PY=

:: Try the Windows py launcher first (most reliable on Windows)
where py >nul 2>&1
if %ERRORLEVEL%==0 (
    for /f "delims=" %%i in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set PY=%%i
    if defined PY goto :check_version
)

:: Fall back to 'python' on PATH
where python >nul 2>&1
if %ERRORLEVEL%==0 (
    for /f "delims=" %%i in ('python -c "import sys; print(sys.executable)" 2^>nul') do set PY=%%i
    if defined PY goto :check_version
)

:: Last resort: probe common install locations
for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
) do (
    if exist %%P (
        set PY=%%~P
        goto :check_version
    )
)

echo  ERROR: Python 3.10 or newer was not found.
echo.
echo  Please install Python from https://www.python.org/downloads/
echo  During installation, tick "Add Python to PATH".
echo.
pause
exit /b 1

:check_version
:: Verify version is 3.10 or higher
for /f "tokens=*" %%v in ('"%PY%" -c "import sys; print(1 if sys.version_info>=(3,10) else 0)" 2^>nul') do set OK=%%v
if "%OK%"=="0" (
    for /f "tokens=*" %%v in ('"%PY%" --version 2^>^&1') do echo  Found: %%v
    echo.
    echo  ERROR: Python 3.10 or newer is required.
    echo  Please upgrade at https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('"%PY%" --version 2^>^&1') do echo  Python : %%v
echo  Path   : %PY%
echo.

:: ------------------------------------------------------------------
:: 2. Create virtual environment
:: ------------------------------------------------------------------
if exist "%VENV_PY%" (
    echo  Virtual environment already exists -- skipping creation.
) else (
    echo  Creating virtual environment in .\venv\ ...
    "%PY%" -m venv "%VENV_DIR%"
    if %ERRORLEVEL% neq 0 (
        echo.
        echo  ERROR: Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo  Virtual environment created.
)
echo.

:: ------------------------------------------------------------------
:: 3. Upgrade pip + install dependencies
:: ------------------------------------------------------------------
echo  Upgrading pip...
"%VENV_PY%" -m pip install --upgrade pip --quiet
echo.

echo  Installing dependencies from requirements.txt...
"%VENV_PY%" -m pip install -r "%APP_DIR%requirements.txt"
if %ERRORLEVEL% neq 0 (
    echo.
    echo  ERROR: Dependency installation failed.
    echo  Check the error messages above, then re-run deploy.bat.
    echo.
    pause
    exit /b 1
)

echo.
echo =====================================================
echo   Deployment complete!
echo.
echo   To start the app:  double-click  launch.bat
echo   To stop the app:   double-click  stop.bat
echo =====================================================
echo.
pause
