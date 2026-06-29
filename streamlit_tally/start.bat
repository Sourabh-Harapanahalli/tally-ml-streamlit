@echo off
REM ==========================================================================
REM Start the TALLY ML Streamlit app on Windows.
REM Run build.bat first to create the virtual environment.
REM ==========================================================================

cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo Virtual environment not found.
    echo Please run build.bat first to set up dependencies.
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat

echo Starting TALLY ML... the app will open in your browser.
streamlit run app.py

pause
