@echo off
echo ==========================================
echo  Nashik iCenter Orderbook Dashboard
echo ==========================================
echo.

REM Check if virtual environment exists, create if not
if not exist ".venv\" (
    echo Creating virtual environment...
    python -m venv .venv
    echo.
)

REM Activate virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate.bat

REM Check if streamlit is installed in venv
.venv\Scripts\pip show streamlit >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing required packages...
    .venv\Scripts\pip install -r requirements.txt
    echo.
)

echo Starting Dashboard...
echo Open your browser at http://localhost:8501
echo Press Ctrl+C to stop the server
echo.
.venv\Scripts\streamlit run app.py --server.port 8501
pause
