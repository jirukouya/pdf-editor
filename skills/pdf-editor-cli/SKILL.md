---
name: pdf-editor-cli
description: Use when an AI needs to run PDF EDITOR from its repository to split a PDF using CSV or XLSX input or merge two PDFs via the interactive CLI, without using the .command launchers.
---

# PDF EDITOR CLI

## When to use this skill

Use this skill when you need to operate PDF EDITOR directly from the repository to either split a merged PDF into smaller PDFs named from a CSV or XLSX sheet, or merge two PDFs into one output PDF.

This skill is for the current interactive CLI only. It is not for MCP, APIs, or the macOS `.command` launchers.

## Current project contract

- Primary entrypoint: `python3 -m pdf_editor`
- Run from the repository root.
- Repository root means the directory that contains `pyproject.toml`, the `pdf_editor/` package, and this `skills/` folder.
- Supported sheet inputs: `.csv`, `.xlsx`
- Supported document input: `.pdf`
- The CLI also supports a fast non-interactive mode through `--mode split` or `--mode merge`.
- The CLI starts by asking whether to run `Split PDF` or `Merge PDF`.
- The CLI auto-detects a name column and may auto-detect an order column.
- The CLI asks for a full naming template that must include `{Name}`.
- If the output folder is left blank, the default folder name follows the fixed text in the naming template. If the template is name-only, it falls back to the source PDF stem.
- Duplicate output filenames are auto-renamed automatically.
- For merge mode, if the output path is blank, the CLI creates a `Merged PDF` folder and uses the first PDF filename by default.
- The CLI writes `split_report.txt` for split mode and `merge_report.txt` for merge mode.
- Startup checks may attempt dependency recovery and may restart inside the local `.venv`.

## How to find the repo root

Do not hardcode any user-specific absolute path.

Find the repository root by using one of these methods:

1. If you are already working inside the repo, use the directory that contains `pyproject.toml` and `pdf_editor/`.
2. If git is available, prefer `git rev-parse --show-toplevel`.
3. If needed, walk upward from this skill file until you find the directory that contains `pyproject.toml` and `pdf_editor/`.

## How to run the CLI

From the repository root, launch:

```bash
python3 -m pdf_editor
```

Normal AI operation should use this command, not `Setup PDF Editor.command` or `Launch PDF Editor.command`.

If you have PTY/stdin control, drive the prompts directly. If you do not have stdin control, guide the human through the same prompt sequence instead of inventing flags or hidden commands.

## Fast CLI mode

Prefer fast CLI mode when all inputs are already known and you do not need the prompt flow.

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
  --first-pdf-path "/path/to/first.pdf" \
  --second-pdf-path "/path/to/second.pdf" \
  --output-path "/path/to/merged.pdf"
```

Fast CLI notes:

- Split mode requires `--sheet-path` and `--pdf-path`
- Merge mode requires `--first-pdf-path` and `--second-pdf-path`
- If `--output-path` is omitted in merge mode, the CLI creates a `Merged PDF` folder and uses the first PDF filename
- If `--output-path` points to an existing folder in merge mode, the CLI uses the first PDF filename inside that folder
- If `--name-column` or `--order-column` is provided in split mode, matching is case-insensitive and spacing-insensitive

## Interactive prompt sequence

Wait for each prompt before sending the next answer. Prefer quoted absolute paths when supplying file paths.

Shared first step:

1. Choose `Split PDF` or `Merge PDF`.

Split flow:

1. Enter the CSV/XLSX path.
2. Review detected sheet columns.
3. Confirm the detected name column, or correct it if detection is wrong.
4. Confirm the detected order column, or leave it blank to use original row order.
5. Enter the PDF path.
6. Enter pages per split PDF.
7. Enter the full naming template and include `{Name}`.
8. Enter the output folder path, or leave it blank for automatic folder creation.
9. Review the summary and confirm yes/no before generation starts.

Merge flow:

1. Enter the first PDF path.
2. Enter the second PDF path.
3. Enter the output PDF path, or leave it blank for automatic output in `Merged PDF`.
4. Review the summary and confirm yes/no before generation starts.

Short prompt-flow sketch:

```text
choose mode -> split flow or merge flow -> summary -> final confirmation
```

## Validation and warnings

Surface validation details to the user before confirming execution.

- If the name column is not detected in split mode, use the CLI correction path and provide one of the detected column names.
- If the order column is not suitable in split mode, either provide a valid detected column name or leave it blank.
- If sheet record count and PDF chunk count do not match in split mode, explicitly show that warning to the user before confirming.
- If duplicate names are reported in split mode, tell the user that output files will be auto-renamed.
- If startup dependency checks fail, let the CLI handle recovery. Do not claim that the dependency is optional.

## Expected outputs

After a successful split run, report:

- the output folder path
- the `split_report.txt` path
- the generated file count
- any unused names or unwritten chunks shown by the CLI

After a successful merge run, report:

- the merged PDF path
- the output folder path
- the `merge_report.txt` path
- the total merged pages

If example output filenames are printed in split mode, you may relay a few of them to confirm naming.

## Do not do these things

- Do not bypass the interactive prompt flow.
- Do not invent nonexistent flags, APIs, or library entrypoints.
- Do not claim MCP support exists in this project.
- Do not use `.command` launchers for normal AI-driven CLI operation.
- Do not auto-confirm warnings without user visibility.
- Do not hardcode or expose user-specific local paths in repo-tracked skill content.
