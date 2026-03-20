# PDF EDITOR

PDF EDITOR is a macOS-focused tool for splitting a merged PDF into smaller PDFs named from a CSV or XLSX file.

It is designed to be simple for non-technical users: after setup, they can start it by double-clicking a launcher file.

## Features

- Interactive step-by-step CLI workflow
- CSV input support
- XLSX input support
- automatic name column detection
- optional order column detection
- filename suffix input
- automatic output folder creation
- duplicate filename auto-renaming
- text report generation
- macOS double-click setup and launch flow

## Recommended Use on macOS

1. Double-click `Setup PDF Editor.command` once
2. Wait for setup to finish
3. Double-click `Launch PDF Editor.command`

This opens Terminal automatically and runs the tool for you.

## First-Time Setup

`Setup PDF Editor.command` will:

- create a local `.venv` virtual environment
- install PDF EDITOR and its required dependencies

After that, most users only need `Launch PDF Editor.command`.

## What the Tool Asks For

1. CSV/XLSX path
2. PDF path
3. pages per split
4. filename suffix after the person's name
5. output folder path, or blank for automatic folder creation

Then it will:

- preview the output filename format
- show a validation summary
- ask for confirmation
- generate the PDFs
- create a report file

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
