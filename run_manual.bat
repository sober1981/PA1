@echo off
REM PA1 - Manual Wednesday Report
REM Double-click this file to run the Wednesday report interactively.
REM You'll be prompted to pick a master file and a week.
REM PDF + Weekly KPI Excel are generated and emailed.

cd /d "C:\Users\jsoberanes\Projects\scorecard-pa"

"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_agent.py --report wednesday

echo.
echo --------------------------------------------------------
echo Press any key to close this window.
pause >nul
