#!/bin/zsh

pause_before_close() {
  if [ -t 0 ]; then
    echo ""
    read "?Press Enter to close..."
  fi
}

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR" || exit 1

clear
echo "=================================================="
echo "PDF EDITOR SETUP"
echo "=================================================="
echo ""

if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 was not found on this Mac."
  echo "Please install Python 3 first, then run this setup again."
  pause_before_close
  exit 1
fi

if [ ! -d ".venv" ]; then
  echo "Creating virtual environment..."
  python3 -m venv .venv
  if [ $? -ne 0 ]; then
    echo ""
    echo "Failed to create the virtual environment."
    pause_before_close
    exit 1
  fi
else
  echo "Virtual environment already exists."
fi

echo ""
echo "Installing PDF EDITOR and required libraries..."
"$SCRIPT_DIR/.venv/bin/python" -m pip install --upgrade pip
"$SCRIPT_DIR/.venv/bin/python" -m pip install -e "$SCRIPT_DIR"

if [ $? -ne 0 ]; then
  echo ""
  echo "Setup failed while installing required libraries."
  pause_before_close
  exit 1
fi

echo ""
echo "Setup completed successfully."
echo "Next time, double-click 'Launch PDF Editor.command' to start the tool."
pause_before_close
