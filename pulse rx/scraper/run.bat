@echo off
REM Launch the Pulse Rx Description Scraper UI.
cd /d "%~dp0"
echo Installing/updating dependencies (first run only)...
python -m pip install --quiet -r requirements.txt
python app.py
pause
