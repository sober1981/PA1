@echo off
REM PA1 - Friday Report (auto-scheduled)
REM Picks the QC'd version of the Wednesday file from Teams sync
REM and sends the report via Outlook.

cd /d "C:\Users\jsoberanes\Projects\scorecard-pa"
"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_agent.py --report friday >> "C:\Users\jsoberanes\Projects\scorecard-pa\logs\friday_%date:~-4%%date:~4,2%%date:~7,2%.log" 2>&1
