@echo off
REM ==========================================================================
REM Build / setup the TALLY ML Streamlit app on Windows.
REM Creates a local virtual environment (.venv) and installs dependencies.
REM Run this once before using start.bat (re-run it to update packages).
REM ==========================================================================

cd /d "%~dp0"

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
