@echo off
REM ==========================================================================
REM Update the Tally AI app to the latest code and (re)start it.
REM
REM   1. If the app is running, close it.
REM   2. Pull the latest code from git ("git pull" on the current branch).
REM   3. If requirements.txt changed, install the new dependencies.
REM   4. Start the app.
REM
REM If the app is not running, steps 2-4 still happen: it syncs and starts.
REM Run build.bat once first to create the virtual environment.
REM ==========================================================================

cd /d "%~dp0"
set "APP_DIR=%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo Virtual environment not found.
    echo Please run build.bat first to set up dependencies.
    pause
    exit /b 1
)

where git >nul 2>nul
if errorlevel 1 (
    echo ERROR: git was not found.
    echo Make sure Git is installed and on your PATH ^(git --version^).
    pause
    exit /b 1
)

REM ---- 1. Close the app if it is running ----------------------------------
REM Stops only processes that belong to THIS folder:
REM   - Streamlit started from this folder's .venv, and
REM   - the start.bat / sync_and_restart.bat window that launched it.
REM Other programs, and this window, are left alone.
REM Exit code 10 = something was stopped, 0 = the app was not running.
echo Checking whether Tally AI is running...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$dir = $env:APP_DIR.TrimEnd('\').ToLower();" ^
  "$venv = $dir + '\.venv\scripts\';" ^
  "$me = (Get-CimInstance Win32_Process -Filter ('ProcessId=' + $PID)).ParentProcessId;" ^
  "$found = Get-CimInstance Win32_Process | Where-Object {" ^
  "  $cmd = ([string]$_.CommandLine).ToLower();" ^
  "  $exe = ([string]$_.ExecutablePath).ToLower();" ^
  "  $isApp = $cmd.Contains($venv + 'streamlit') -or $exe.Contains($venv + 'streamlit') -or ($exe.Contains($venv) -and $cmd.Contains('-m streamlit'));" ^
  "  $isWindow = $_.Name -eq 'cmd.exe' -and ($cmd.Contains($dir + '\start.bat') -or $cmd.Contains($dir + '\sync_and_restart.bat'));" ^
  "  $_.ProcessId -ne $PID -and $_.ProcessId -ne $me -and ($isApp -or $isWindow)" ^
  "};" ^
  "if ($found) { $found | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }; exit 10 };" ^
  "exit 0"
set "STOP_RESULT=%errorlevel%"
if "%STOP_RESULT%"=="10" (
    echo Tally AI was running - closed it.
    REM Give Windows a moment to release the port the app was using.
    timeout /t 2 /nobreak >nul
)
if "%STOP_RESULT%"=="0" echo Tally AI was not running.
if not "%STOP_RESULT%"=="10" if not "%STOP_RESULT%"=="0" (
    echo WARNING: could not check whether Tally AI is running.
    echo If it is, close its window yourself, then run this file again.
)

REM ---- 2. Sync ------------------------------------------------------------
set "REQ_BEFORE="
for /f "delims=" %%H in ('git hash-object requirements.txt 2^>nul') do set "REQ_BEFORE=%%H"

echo.
echo Pulling the latest changes from git...
set "SYNC_OK=1"
git pull
if errorlevel 1 set "SYNC_OK=0"

if "%SYNC_OK%"=="0" (
    echo.
    echo WARNING: git pull failed, so the app was NOT updated.
    echo Check your internet connection, and make sure you have no local
    echo changes that conflict with the incoming update.
    echo Starting the app with the code already on this computer...
    timeout /t 8
)

REM ---- 3. Install dependencies if requirements.txt changed ----------------
set "REQ_AFTER="
for /f "delims=" %%H in ('git hash-object requirements.txt 2^>nul') do set "REQ_AFTER=%%H"
set "REQ_CHANGED=0"
if not "%REQ_BEFORE%"=="%REQ_AFTER%" set "REQ_CHANGED=1"

call .venv\Scripts\activate.bat

if "%REQ_CHANGED%"=="1" (
    echo.
    echo requirements.txt changed - installing the updated dependencies...
    pip install -r requirements.txt
)
if "%REQ_CHANGED%"=="1" if errorlevel 1 (
    echo.
    echo WARNING: installing dependencies failed. If the app does not
    echo start, run build.bat and then this file again.
    timeout /t 8
)

REM ---- 4. Start -----------------------------------------------------------
echo.
echo Starting Tally AI... the app will open in your browser.
streamlit run app.py

pause
