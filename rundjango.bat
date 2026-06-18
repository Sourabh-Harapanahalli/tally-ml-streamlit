@echo off

REM Replace this path with the path to your Django project folder
cd C:\Users\554so\OneDrive\Desktop\TALLY_ML\TALLY ML

REM Activate the virtual environment (if applicable)
call venv\Scripts\activate

REM Change directory to your Django project folder (if needed)
cd GST

REM Start Django development server in the background
start "" python manage.py runserver

REM Wait for a few seconds to allow the server to start
timeout /t 5

REM Open the Django application in the default web browser
start http://127.0.0.1:8000/

REM Pause to keep the command prompt window open (optional)
pause
