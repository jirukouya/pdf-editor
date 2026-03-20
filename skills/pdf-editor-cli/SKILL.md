---
name: pdf-editor-cli
description: Use when an AI needs to run PDF EDITOR from its repository to split a PDF using CSV or XLSX input via the interactive CLI, without using the .command launchers.
---

# PDF EDITOR CLI

## When to use this skill

Use this skill when you need to operate PDF EDITOR directly from the repository to split a merged PDF into smaller PDFs named from a CSV or XLSX sheet.

This skill is for the current interactive CLI only. It is not for MCP, APIs, or the macOS `.command` launchers.

## Current project contract

- Primary entrypoint: `python3 -m pdf_editor`
- Run from the repository root.
- Repository root means the directory that contains `pyproject.toml`, the `pdf_editor/` package, and this `skills/` folder.
- Supported sheet inputs: `.csv`, `.xlsx`
- Supported document input: `.pdf`
- The CLI is interactive. There is no documented non-interactive split command with flags.
- The CLI auto-detects a name column and may auto-detect an order column.
- If the output folder is left blank, the default folder name follows the filename suffix. If the suffix is blank, it falls back to the source PDF stem.
- Duplicate output filenames are auto-renamed automatically.
- The CLI writes `split_report.txt` into the output folder.
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

## Interactive prompt sequence

Wait for each prompt before sending the next answer. Prefer quoted absolute paths when supplying file paths.

1. Enter the CSV/XLSX path.
2. Review detected sheet columns.
3. Confirm the detected name column, or correct it if detection is wrong.
4. Confirm the detected order column, or leave it blank to use original row order.
5. Enter the PDF path.
6. Enter pages per split PDF.
7. Enter the filename suffix after the person's name, or leave it blank.
8. Enter the output folder path, or leave it blank for automatic folder creation.
9. Review the summary and confirm yes/no before generation starts.

Short prompt-flow sketch:

```text
sheet path -> column review -> name/order confirmation -> PDF path -> pages per file -> suffix -> output folder -> summary -> final confirmation
```

## Validation and warnings

Surface validation details to the user before confirming execution.

- If the name column is not detected, use the CLI correction path and provide one of the detected column names.
- If the order column is not suitable, either provide a valid detected column name or leave it blank.
- If sheet record count and PDF chunk count do not match, explicitly show that warning to the user before confirming.
- If duplicate names are reported, tell the user that output files will be auto-renamed.
- If startup dependency checks fail, let the CLI handle recovery. Do not claim that the dependency is optional.

## Expected outputs

After a successful run, report:

- the output folder path
- the `split_report.txt` path
- the generated file count
- any unused names or unwritten chunks shown by the CLI

If example output filenames are printed, you may relay a few of them to confirm naming.

## Do not do these things

- Do not bypass the interactive prompt flow.
- Do not invent nonexistent flags, APIs, or library entrypoints.
- Do not claim MCP support exists in this project.
- Do not use `.command` launchers for normal AI-driven CLI operation.
- Do not auto-confirm warnings without user visibility.
- Do not hardcode or expose user-specific local paths in repo-tracked skill content.
