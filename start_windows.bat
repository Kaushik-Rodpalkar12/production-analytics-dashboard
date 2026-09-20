@echo off
cd /d "%~dp0"
echo Installing packages (first run only)...
pip install -r requirements.txt
echo.
echo Starting the dashboard...
streamlit run app.py
pause
