@echo off
cd /d "%~dp0"
echo =========================================
echo   PASTE YOUR GOOGLE API KEY BELOW
echo =========================================
set /p GOOGLE_PLACES_API_KEY="API Key: "
echo Running script...
python process_restaurants.py
pause