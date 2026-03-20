# PDF EDITOR v0.1.0

PDF EDITOR is a macOS-focused PDF splitting tool built around an interactive CLI workflow.

## Highlights

- Interactive step-by-step CLI workflow
- CSV and XLSX input support
- macOS double-click setup and launch files
- automatic filename suffix handling
- automatic output folder creation
- duplicate filename auto-renaming
- text report generation
- Python 3.9+ compatibility
- local `.venv` fallback when dependency installation cannot use the system Python environment

## Download and Use

1. Download `PDF Editor.zip` from this release
2. Extract the zip file
3. Double-click `Setup PDF Editor.command`
4. After setup finishes, double-click `Launch PDF Editor.command`

## Notes

- macOS may ask you to confirm opening `.command` files the first time.
- The setup script creates a local virtual environment and installs PDF EDITOR with its dependencies.
- This release is intended for macOS.
- MCP server support and Agent Skills are not bundled yet; the current repository is CLI-first and ready to be wrapped by those layers later.
