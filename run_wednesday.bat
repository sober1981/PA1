@echo off
REM PA1 - Wednesday Report (auto-scheduled safety net)
REM Runs at 14:00 on Wednesdays. Skips if the report was already generated
REM manually earlier in the day (checks state/last_run.json).
REM If the report fails, a failure notification email is sent.

cd /d "C:\Users\jsoberanes\Projects\scorecard-pa"

set LOGFILE=C:\Users\jsoberanes\Projects\scorecard-pa\logs\wednesday_%date:~-4%%date:~4,2%%date:~7,2%.log

"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_agent.py --report wednesday >> "%LOGFILE%" 2>&1

if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] Report failed with exit code %ERRORLEVEL% >> "%LOGFILE%"
    "C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" -c "from src.emailer import send_report_email; from datetime import datetime; log=open(r'%LOGFILE%','r').read()[-2000:]; send_report_email(subject='PA1 - REPORT FAILED (%s)' %% datetime.now().strftime('%%Y-%%m-%%d %%H:%%M'), body_text='PA1 Wednesday report failed.\n\nLast 2000 chars of log:\n' + log + '\n\n-- PA1', pdf_paths=[], recipient='jsoberanes@scoutdownhole.com')" 2>> "%LOGFILE%"
)
