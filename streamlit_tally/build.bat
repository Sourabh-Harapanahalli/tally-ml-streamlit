@echo off
REM ==========================================================================
REM Build / setup the TALLY ML Streamlit app on Windows.
REM Always does a clean install: removes any existing .venv, recreates it,
REM and installs dependencies from scratch.
REM ==========================================================================

cd /d "%~dp0"

if exist ".venv" (
    echo Removing existing virtual environment for a clean rebuild...
    rmdir /s /q ".venv"
)

echo Creating virtual environment (.venv)...
python -m venv .venv
if errorlevel 1 (
    echo.
    echo ERROR: could not create the virtual environment.
    echo Make sure Python 3.11+ is installed and on your PATH ^(python --version^).
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat

echo Upgrading pip...
python -m pip install --upgrade pip

echo Installing dependencies from requirements.txt...
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: dependency installation failed.
    pause
    exit /b 1
)

echo.
echo Build complete. Run start.bat to launch the app.
pause
