@echo off
REM Automated processing run for accounting mailbox reader
REM Called by Windows Task Scheduler hourly 08:00-17:00
REM At 17:00 (last run), also runs Reminders cleanup with remaining daily budget

cd /d "c:\Users\MatthiasDiBianco\accounting-mailbox-reader"
call .venv\Scripts\activate.bat

REM Log rotation: clear log if older than 30 days
set LOG=%USERPROFILE%\.accounting_mailbox_reader\process.log
powershell -Command "if (Test-Path '%LOG%') { if ((Get-Date) - (Get-Item '%LOG%').LastWriteTime -gt [TimeSpan]::FromDays(30)) { Clear-Content '%LOG%'; Add-Content '%LOG%' '[Log cleared after 30 days]' } }"

echo. >> "%LOG%"
echo [%date% %time%] === Scheduled run starting === >> "%LOG%"

REM Normal process run (exit code 1 = health check failed, skip cleanup)
python main.py process --upload-sharepoint >> "%LOG%" 2>&1
set PROCESS_EXIT=%ERRORLEVEL%

if %PROCESS_EXIT% NEQ 0 (
    echo [%date% %time%] Process failed with exit code %PROCESS_EXIT%, skipping cleanup >> "%LOG%"
    echo [%date% %time%] === Run failed === >> "%LOG%"
    exit /b %PROCESS_EXIT%
)

REM DISABLED: cleanup deactivated while Jira integration is active
REM for /f "tokens=1 delims=:" %%h in ("%time: =0%") do set hour=%%h
REM if %%hour%% GEQ 17 (
REM     echo [%date% %time%] Running Reminders cleanup with remaining budget... >> "%LOG%"
REM     python main.py cleanup --upload-sharepoint >> "%LOG%" 2>&1
REM )

echo [%date% %time%] === Run complete === >> "%LOG%"
