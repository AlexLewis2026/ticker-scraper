@echo off
echo Starting Trade Accumulator...
start "" http://localhost:5001
"C:\Users\AlexLewis\AppData\Local\Python\bin\python3.exe" "%~dp0app.py"
pause
