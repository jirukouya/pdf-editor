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
- validation-first and warning handling
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

Fast CLI safe-mode flags:

- `--dry-run` or `--validate-only`: validate only, never write files
- `--json`: print one JSON object to stdout for automation
- `--confirm`: allow execution to continue when validation returns warnings
- `--strict`: treat warnings as errors and stop
- `--on-output-exists fail|overwrite|rename|continue`: explicit output conflict policy
- `--duplicate-name-policy autorename|fail|append-row-number|append-order`: split duplicate rendered filename policy

Fast CLI exit codes:

- `0`: validation passed or execution succeeded
- `2`: warning state, execution blocked until `--confirm`
- `1`: hard error or strict-mode rejection

Recommended automation pattern:

1. Run validation first with `--dry-run --json`
2. Inspect `status`, `warnings`, `errors`, and `requires_confirmation`
3. Only rerun without `--dry-run` when status is safe or after explicit approval with `--confirm`

Split validation example:

```bash
python3 -m pdf_editor \
  --mode split \
  --sheet-path "/path/to/employees.xlsx" \
  --pdf-path "/path/to/source.pdf" \
  --pages-per-file 1 \
  --naming-template "GD Pink Form - Letter of Offer ({Name}) 26-3-2026" \
  --output-dir "/path/to/output" \
  --dry-run \
  --json
```

Split execution example:

```bash
python3 -m pdf_editor \
  --mode split \
  --sheet-path "/path/to/employees.xlsx" \
  --pdf-path "/path/to/source.pdf" \
  --pages-per-file 1 \
  --naming-template "GD Pink Form - Letter of Offer ({Name}) 26-3-2026" \
  --output-dir "/path/to/output" \
  --on-output-exists rename \
  --duplicate-name-policy append-order \
  --confirm
```

Simple merge validation example:

```bash
python3 -m pdf_editor \
  --mode merge \
  --merge-kind simple \
  --first-pdf-path "/path/to/first.pdf" \
  --second-pdf-path "/path/to/second.pdf" \
  --output-path "/path/to/merged.pdf" \
  --dry-run \
  --json
```

Simple merge execution example:

```bash
python3 -m pdf_editor \
  --mode merge \
  --merge-kind simple \
  --first-pdf-path "/path/to/first.pdf" \
  --second-pdf-path "/path/to/second.pdf" \
  --output-path "/path/to/merged.pdf" \
  --on-output-exists overwrite \
  --confirm
```

Batch merge validation example:

```bash
python3 -m pdf_editor \
  --mode merge \
  --merge-kind batch \
  --batch-input-dir "/path/to/split-output" \
  --fixed-pdf-path "/path/to/fixed.pdf" \
  --merge-order split-first \
  --batch-output-dir "/path/to/batch-output" \
  --dry-run \
  --json
```

Batch merge execution example:

```bash
python3 -m pdf_editor \
  --mode merge \
  --merge-kind batch \
  --batch-input-dir "/path/to/split-output" \
  --fixed-pdf-path "/path/to/fixed.pdf" \
  --merge-order split-first \
  --batch-output-dir "/path/to/batch-output" \
  --on-output-exists rename \
  --confirm
```

Fast CLI defaults:

- Split mode defaults `--pages-per-file` to `1`
- Split mode defaults `--naming-template` to `{Name}`
- Simple merge creates a `Merged PDF` folder automatically if `--output-path` is omitted
- Simple merge uses the first PDF filename automatically if `--output-path` points to a folder or is omitted
- Batch merge creates a `Batch Merged PDF` folder automatically if `--batch-output-dir` is omitted
- Batch merge keeps each split PDF filename by default and only adds `(2)` when needed
- Warnings block execution by default until the same command is rerun with `--confirm`
- `--strict` upgrades warnings into hard failures
- `--json` keeps stdout machine-readable and moves human-readable warning text out of stdout
- `--on-output-exists` defaults to `fail` for explicit fast-CLI outputs
- `--duplicate-name-policy` defaults to `autorename` in split fast CLI
- `rename` and `continue` both preserve existing files and generate new unique paths when needed

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
