@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt
pause
echo Setup complete! You can now double-click project_time_tracker.py to run the app.
pause