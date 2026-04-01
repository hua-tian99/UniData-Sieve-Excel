@echo off
chcp 65001 >nul
echo Starting UniData-Sieve Web Interface...

REM 优先尝试激活 .venv，其次尝试 venv 目录
if exist ".venv\Scripts\activate.bat" (
    echo Activating .venv virtual environment...
    call ".venv\Scripts\activate.bat"
) else if exist "venv\Scripts\activate.bat" (
    echo Activating venv virtual environment...
    call "venv\Scripts\activate.bat"
) else (
    echo WARNING: Virtual environment not found. Using system Python...
)

echo Starting Streamlit...
streamlit run app.py
pause
