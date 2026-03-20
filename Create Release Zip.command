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
echo "PDF EDITOR RELEASE ZIP"
echo "=================================================="
echo ""

mkdir -p dist
rm -f dist/'PDF Editor.zip'

zip -r dist/'PDF Editor.zip'   'Setup PDF Editor.command'   'Launch PDF Editor.command'   'README.md'   'RELEASE_NOTES_v0.1.0.md'   'pyproject.toml'   'pdf_editor'   -x '*/__pycache__/*' '*.pyc' '.DS_Store'

if [ $? -ne 0 ]; then
  echo ""
  echo "Failed to create release zip."
  pause_before_close
  exit 1
fi

echo ""
echo "Release package created:"
echo "$SCRIPT_DIR/dist/PDF Editor.zip"
pause_before_close
