@echo off
chcp 65001 >nul
echo Building UniData-Sieve Standalone Package...
echo Please wait 1-3 minutes for the first PyInstaller execution.

call .venv\Scripts\activate

pyinstaller --noconfirm --onedir --clean --name "UniData-Sieve" --add-data "app.py;." --add-data "config;config" --add-data "engine;engine" --copy-metadata streamlit --copy-metadata loguru --copy-metadata pydantic --collect-all streamlit --collect-all openpyxl --hidden-import pandas --hidden-import loguru --hidden-import pydantic --hidden-import streamlit.web.cli --hidden-import pyarrow.vendored.version launcher.py

echo ======================================================
echo Build Successfully Completed!
echo You can find the executable in the /dist/UniData-Sieve folder.
echo ======================================================
pause
