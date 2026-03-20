# PDF EDITOR

PDF EDITOR is a macOS-focused tool for splitting a merged PDF into smaller PDFs named from a CSV or XLSX file.

It supports both:

- a human-friendly macOS launcher flow for non-technical users
- a repo-local Agent Skill for AI-driven CLI operation

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
- repo-local `Skill.MD` for AI agents that should drive the CLI directly

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

## Agent Skill

This repository includes a repo-local skill at [`skills/pdf-editor-cli/SKILL.md`](skills/pdf-editor-cli/SKILL.md).

That skill is intended for AI agents that should operate the project directly through:

```bash
python3 -m pdf_editor
```

The skill documents:

- how to find the repository root without hardcoded personal paths
- the current interactive CLI prompt order
- validation and warning handling
- expected outputs such as `split_report.txt`

The current repository is CLI-first. MCP support is not implemented yet.

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
cd /path/to/pdf-editor-repo
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
- The repo-local Agent Skill is for AI-driven CLI usage and is not bundled into the end-user release zip.
- MCP support is still under consideration.

## License

This project is released under the MIT License. See `LICENSE` for details.
