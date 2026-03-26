# PDF EDITOR

PDF EDITOR is a macOS-focused tool for splitting PDFs from CSV/XLSX naming data and merging PDFs in either simple or batch workflows.

It supports both:

- a human-friendly macOS launcher flow for non-technical users
- a repo-local Agent Skill for AI-driven CLI operation

## Features

- Interactive step-by-step CLI workflow
- CSV input support
- XLSX input support
- automatic name column detection
- optional order column detection
- custom naming template input with `{Name}` placeholder
- simple merge for combining two PDFs into one output file
- batch merge for looping a split-output folder against one fixed PDF
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
- the fast non-interactive CLI flags for AI-friendly runs
- validation and warning handling
- expected outputs such as `split_report.txt` and `merge_report.txt`

The current repository is CLI-first. MCP support is not implemented yet.

## What the Tool Asks For

When you choose `Split PDF`:

1. CSV/XLSX path
2. PDF path
3. pages per split
4. full naming template that includes `{Name}`
5. output folder path, or blank for automatic folder creation

When you choose `Merge PDF`:

Choose either:

- `Simple Merge`
  1. first PDF path
  2. second PDF path
  3. output PDF path, or blank for automatic folder creation in `Merged PDF`
- `Batch Merge`
  1. split-output folder path
  2. fixed PDF path
  3. merge order: split-first or fixed-first
  4. output folder path, or blank for automatic folder creation in `Batch Merged PDF`

Then it will:

- show a preview or summary
- ask for confirmation
- generate the PDF output
- create a report file

## Manual Run

```bash
cd /path/to/pdf-editor-repo
python3 -m pdf_editor
```

## Fast CLI For AI

When an AI or automation already knows all required inputs, it can skip the prompt flow and run the CLI directly with flags.

Split example:

```bash
python3 -m pdf_editor \
  --mode split \
  --sheet-path "/path/to/employees.xlsx" \
  --pdf-path "/path/to/source.pdf" \
  --pages-per-file 1 \
  --naming-template "GD Pink Form - Letter of Offer ({Name}) 26-3-2026" \
  --output-dir "/path/to/output"
```

Merge example:

```bash
python3 -m pdf_editor \
  --mode merge \
  --merge-kind simple \
  --first-pdf-path "/path/to/first.pdf" \
  --second-pdf-path "/path/to/second.pdf" \
  --output-path "/path/to/merged.pdf"
```

Batch merge example:

```bash
python3 -m pdf_editor \
  --mode merge \
  --merge-kind batch \
  --batch-input-dir "/path/to/split-output" \
  --fixed-pdf-path "/path/to/fixed.pdf" \
  --merge-order split-first \
  --batch-output-dir "/path/to/batch-output"
```

Fast CLI defaults:

- Split mode defaults `--pages-per-file` to `1`
- Split mode defaults `--naming-template` to `{Name}`
- Simple merge creates a `Merged PDF` folder automatically if `--output-path` is omitted
- Simple merge uses the first PDF filename automatically if `--output-path` points to a folder or is omitted
- Batch merge creates a `Batch Merged PDF` folder automatically if `--batch-output-dir` is omitted
- Batch merge keeps each split PDF filename by default and only adds `(2)` when needed

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
