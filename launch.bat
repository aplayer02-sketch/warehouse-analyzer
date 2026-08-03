@echo off
title Warehouse Tool
call "%~dp0venv\Scripts\activate.bat"
start http://localhost:8501
python -m streamlit run "%~dp0warehouse_analyzer.py"
pause
