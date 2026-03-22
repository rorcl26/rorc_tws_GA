@echo off
cd /d %~dp0
echo Installing requirements...
python -m pip install --user -r requirements.txt
echo Running analysis...
python stock_analyzer.py
pause
