@echo off
REM Accounting Mailbox Reader - Windows Helper Script

cd /d "%~dp0"

if "%1"=="" (
    echo.
    echo Usage: run.bat [command]
    echo.
    echo Commands:
    echo   setup        - Initialize environment and install dependencies
    echo   config       - Show configuration
    echo   read         - Read emails (usage: run.bat read [options])
    echo   preview      - Quick preview of recent emails
    echo   shell        - Activate virtual environment
    echo.
    echo Examples:
    echo   run.bat config
    echo   run.bat preview
    echo   run.bat read --format json --output emails.json
    echo.
    exit /b 0
)

if "%1"=="setup" (
    echo Creating virtual environment...
    python -m venv .venv
    echo Installing dependencies...
    .\.venv\Scripts\pip install -r requirements.txt
    echo Initializing .env file...
    .\.venv\Scripts\python main.py init
    echo.
    echo Setup complete! Run: run.bat config
    exit /b 0
)

if "%1"=="config" (
    .\.venv\Scripts\python main.py config-show
    exit /b 0
)

if "%1"=="preview" (
    .\.venv\Scripts\python main.py preview %2 %3 %4 %5
    exit /b 0
)

if "%1"=="read" (
    .\.venv\Scripts\python main.py read %2 %3 %4 %5 %6 %7
    exit /b 0
)

if "%1"=="shell" (
    echo Activating virtual environment...
    .\.venv\Scripts\activate.bat
    exit /b 0
)

echo Unknown command: %1
exit /b 1
