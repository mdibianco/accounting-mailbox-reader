@echo off
REM Automated processing run for accounting mailbox reader
REM This script is called by Windows Task Scheduler

cd /d "c:\Users\MatthiasDiBianco\accounting-mailbox-reader"
call .venv\Scripts\activate.bat
python main.py process --upload-sharepoint >> "%USERPROFILE%\.accounting_mailbox_reader\process.log" 2>&1
