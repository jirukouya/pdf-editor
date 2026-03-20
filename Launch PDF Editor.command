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

if [ ! -d ".venv" ]; then
  echo "PDF EDITOR is not set up yet."
  echo "Please double-click 'Setup PDF Editor.command' first."
  pause_before_close
  exit 1
fi

"$SCRIPT_DIR/.venv/bin/python" -m pdf_editor
STATUS=$?

echo ""
if [ $STATUS -eq 0 ]; then
  echo "PDF EDITOR finished."
elif [ $STATUS -eq 130 ]; then
  echo "PDF EDITOR was cancelled."
else
  echo "PDF EDITOR exited with code $STATUS."
fi
pause_before_close
exit $STATUS
