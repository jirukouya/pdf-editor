# PDF EDITOR

PDF EDITOR is an interactive macOS tool for splitting a merged PDF into smaller PDFs named from a CSV or XLSX file.

It is designed for users who do not want to work directly in Terminal. After setup, the tool can be launched by double-clicking a file.

## Features

- Interactive step-by-step workflow
- Supports CSV and XLSX input files
- Auto-detects a name column and an optional order column
- Lets the user enter a filename suffix for output PDFs
- Auto-creates an output folder when left blank
- Default output folder follows the user-provided filename suffix
- Auto-renames files when name conflicts already exist
- Generates a text report after each run

## Recommended Use on macOS

For non-technical users:

1. Double-click `Setup PDF Editor.command` once
2. Wait for setup to finish
3. Double-click `Launch PDF Editor.command`

This opens Terminal automatically and runs the tool for you.

## First-Time Setup

`Setup PDF Editor.command` will:

- create a local `.venv` virtual environment
- install the required dependency `pypdf`

After that, users normally only need to open `Launch PDF Editor.command`.

## End-User Flow

The tool will ask for:

1. CSV/XLSX path
2. PDF path
3. Pages per split
4. Filename suffix after the person's name
5. Output folder path, or blank for automatic folder creation

Then it will:

- preview the output filename format
- show a validation summary
- ask for confirmation
- generate the PDFs
- create a report file

## Repository Contents

Keep these in the repository:

- `pdf_editor/`
- `tests/`
- `Setup PDF Editor.command`
- `Launch PDF Editor.command`
- `Create Release Zip.command`
- `README.md`
- `pyproject.toml`
- `.gitignore`

Do not commit:

- `.venv/`
- `__pycache__/`
- `*.pyc`
- generated output PDF folders
- local release zip files in `dist/`

## Create a GitHub Release Zip

Double-click:

- `Create Release Zip.command`

This creates:

- `dist/PDF Editor.zip`

That zip file is the one you can upload to GitHub Releases for end users.

## Manual Run

```bash
cd '/Users/derekho/Documents/Coding/Python/PDF Editor'
python3 -m pdf_editor
```

## Optional Editable Install

```bash
python3 -m pip install -e .
pdf-editor
```

## Notes

- The project currently targets macOS.
- XLSX support does not require `openpyxl`.
- If users download the project from GitHub, macOS may ask them to confirm opening `.command` files the first time.

## License

This project is released under the MIT License. See `LICENSE` for details.
