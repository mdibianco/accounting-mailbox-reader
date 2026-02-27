#!/bin/bash
# Accounting Mailbox Reader - macOS/Linux Helper Script

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

if [ -z "$1" ]; then
    echo ""
    echo "Usage: ./run.sh [command]"
    echo ""
    echo "Commands:"
    echo "  setup        - Initialize environment and install dependencies"
    echo "  config       - Show configuration"
    echo "  read         - Read emails (usage: ./run.sh read [options])"
    echo "  preview      - Quick preview of recent emails"
    echo "  shell        - Activate virtual environment"
    echo ""
    echo "Examples:"
    echo "  ./run.sh config"
    echo "  ./run.sh preview"
    echo "  ./run.sh read --format json --output emails.json"
    echo ""
    exit 0
fi

if [ "$1" == "setup" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
    echo "Installing dependencies..."
    ./venv/bin/pip install -r requirements.txt
    echo "Initializing .env file..."
    ./venv/bin/python main.py init
    echo ""
    echo "Setup complete! Run: ./run.sh config"
    exit 0
fi

if [ "$1" == "config" ]; then
    ./venv/bin/python main.py config-show
    exit 0
fi

if [ "$1" == "preview" ]; then
    ./venv/bin/python main.py preview "${@:2}"
    exit 0
fi

if [ "$1" == "read" ]; then
    ./venv/bin/python main.py read "${@:2}"
    exit 0
fi

if [ "$1" == "shell" ]; then
    echo "Activating virtual environment..."
    source venv/bin/activate
    exit 0
fi

echo "Unknown command: $1"
exit 1
