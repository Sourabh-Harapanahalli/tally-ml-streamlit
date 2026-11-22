@echo off
REM ==========================================================================
REM Sync the Tally AI Streamlit app with the latest code from git.
REM Runs "git pull" for the current branch. If requirements.txt changed,
REM run build.bat afterwards to reinstall dependencies.
REM ==========================================================================

cd /d "%~dp0"

where git >nul 2>nul
if errorlevel 1 (
    echo ERROR: git was not found.
    echo Make sure Git is installed and on your PATH ^(git --version^).
    pause
    exit /b 1
)

echo Pulling the latest changes from git...
git pull
if errorlevel 1 (
    echo.
    echo ERROR: git pull failed.
    echo Check your internet connection, and make sure you have no local
    echo changes that conflict with the incoming update.
    pause
    exit /b 1
)

echo.
echo Sync complete. Run start.bat to launch the app.
echo If requirements.txt changed, run build.bat first.
pause
