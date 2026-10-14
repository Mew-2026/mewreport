@echo off
rem Start the Loading Sheet Check web app on this computer: http://127.0.0.1:5000
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python is not installed. Install Python 3.10 or newer from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during setup, then run this file again.
    pause
    exit /b 1
)

if not exist ".env" (
    copy ".env.example" ".env" >nul
    echo Created .env - open it in Notepad, paste your renamed.to API key after RENAMED_API_KEY=
    echo then run this file again.
    notepad ".env"
    pause
    exit /b 1
)

findstr /r /c:"^RENAMED_API_KEY=rt_" ".env" >nul
if errorlevel 1 (
    echo .env has no renamed.to key yet. Paste it after RENAMED_API_KEY= and save.
    notepad ".env"
    pause
    exit /b 1
)

echo Installing / checking required Python packages...
python -m pip install --quiet --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
    echo Package install failed - check the internet connection and try again.
    pause
    exit /b 1
)

echo.
echo Starting the app at http://127.0.0.1:5000  (close this window to stop it)
start "" http://127.0.0.1:5000
python -m flask --app api/index.py run --port 5000
pause
