@echo off
cd /d "%~dp0"
set "PYTHON_CMD=C:\Users\dark_\AppData\Local\Python\pythoncore-3.14-64\python.exe"
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
set "STREAMLIT_SERVER_HEADLESS=true"
if exist "%PYTHON_CMD%" (
  "%PYTHON_CMD%" run_app.py
) else (
  python run_app.py
)
