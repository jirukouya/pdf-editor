---
name: pdf-editor-cli
description: Use when an AI needs to run PDF EDITOR from its repository to split a PDF using CSV or XLSX input or perform simple or batch PDF merge workflows via the CLI, without using the .command launchers.
---

# PDF EDITOR CLI

## When to use this skill

Use this skill when you need to operate PDF EDITOR directly from the repository to either split a merged PDF into smaller PDFs named from a CSV or XLSX sheet, run a simple two-file merge, or run a batch merge over a split-output folder with one fixed PDF.

This skill is for the repository CLI, including both the interactive prompt flow and the fast non-interactive mode. It is not for MCP, APIs, or the macOS `.command` launchers.

## Current project contract

- Primary entrypoint: `python3 -m pdf_editor`
- Run from the repository root.
- Repository root means the directory that contains `pyproject.toml`, the `pdf_editor/` package, and this `skills/` folder.
- Supported sheet inputs: `.csv`, `.xlsx`
- Supported document input: `.pdf`
- The CLI also supports a fast non-interactive mode through `--mode split` or `--mode merge`.
- Fast CLI supports a safe validation-first workflow through `--dry-run`, `--validate-only`, `--json`, `--confirm`, and `--strict`.
- Fast CLI also supports explicit conflict policies through `--on-output-exists` and `--duplicate-name-policy`.
- The CLI starts by asking whether to run `Split PDF` or `Merge PDF`.
- Merge mode then asks whether to run `Simple Merge` or `Batch Merge`.
- The CLI auto-detects a name column and may auto-detect an order column.
- The CLI asks for a full naming template that must include `{Name}`.
- If the output folder is left blank, the default folder name follows the fixed text in the naming template. If the template is name-only, it falls back to the source PDF stem.
- Duplicate output filenames are auto-renamed automatically.
- For simple merge mode, if the output path is blank, the CLI creates a `Merged PDF` folder and uses the first PDF filename by default.
- For batch merge mode, if the output folder is blank, the CLI creates a `Batch Merged PDF` folder.
- For batch merge mode, the output files keep the split PDF filenames by default.
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

Safe fast-mode contract:

- Run validation first with `--dry-run` or `--validate-only`
- Add `--json` when an agent or script needs one machine-readable stdout payload
- If validation returns warnings, execution is blocked unless you rerun with `--confirm`
- If `--strict` is present, warnings become errors and execution stays blocked
- `--on-output-exists fail|overwrite|rename|continue` controls explicit output conflicts
- `--duplicate-name-policy autorename|fail|append-row-number|append-order` controls duplicate rendered filenames in split mode
- Fast CLI exit codes:
  - `0` = validation passed or execution succeeded
  - `2` = warning state, blocked pending confirmation
  - `1` = hard error or strict-mode rejection

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

Fast CLI notes:

- Split mode requires `--sheet-path` and `--pdf-path`
- Simple merge mode requires `--first-pdf-path` and `--second-pdf-path`
- Batch merge mode requires `--batch-input-dir` and `--fixed-pdf-path`
- `--dry-run` and `--validate-only` are equivalent
- `--json` should be preferred for agent-driven automation because stdout becomes a single JSON object
- Warning states require an explicit rerun with `--confirm`
- `--strict` upgrades warnings into errors and overrides `--confirm`
- `--on-output-exists` defaults to `fail`
- `--duplicate-name-policy` defaults to `autorename`
- `rename` and `continue` both preserve existing files and move new outputs onto unique `(2)` style paths
- If `--output-path` is omitted in simple merge mode, the CLI creates a `Merged PDF` folder and uses the first PDF filename
- If `--output-path` points to an existing folder in simple merge mode, the CLI uses the first PDF filename inside that folder
- If `--batch-output-dir` is omitted in batch merge mode, the CLI creates a `Batch Merged PDF` folder
- If `--merge-order` is used in batch merge mode, valid values are `split-first` and `fixed-first`
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

1. Choose `Simple Merge` or `Batch Merge`.

Simple merge flow:

1. Enter the first PDF path.
2. Enter the second PDF path.
3. Enter the output PDF path, or leave it blank for automatic output in `Merged PDF`.
4. Review the summary and confirm yes/no before generation starts.

Batch merge flow:

1. Enter the split-output folder path.
2. Enter the fixed PDF path.
3. Choose the merge order: `split-first` or `fixed-first`.
4. Enter the output folder path, or leave it blank for automatic output in `Batch Merged PDF`.
5. Review the summary and confirm yes/no before generation starts.

Short prompt-flow sketch:

```text
choose mode -> split flow or merge flow -> summary -> final confirmation
```

## Validation and warnings

Surface validation details to the user before confirming execution.

In fast CLI mode, prefer this pattern:

1. Run `--dry-run --json`
2. Inspect `status`, `warnings`, `errors`, `requires_confirmation`, and `summary`
3. Only rerun for execution when validation is clean or the user explicitly approves `--confirm`

- If the name column is not detected in split mode, use the CLI correction path and provide one of the detected column names.
- If the order column is not suitable in split mode, either provide a valid detected column name or leave it blank.
- If sheet record count and PDF chunk count do not match in split mode, explicitly show that warning to the user before confirming.
- If duplicate names are reported in split mode, tell the user that output files will be auto-renamed.
- If duplicate rendered filenames in split mode must be treated as blocking, use `--duplicate-name-policy fail`.
- If an explicit split output directory already exists and is non-empty, the default fast-CLI policy is `--on-output-exists fail`.
- If a simple merge output file already exists, the default fast-CLI policy is `--on-output-exists fail`.
- If an explicit batch merge output directory already exists and is non-empty, the default fast-CLI policy is `--on-output-exists fail`.
- If startup dependency checks fail, let the CLI handle recovery. Do not claim that the dependency is optional.

## Expected outputs

After a successful split run, report:

- the output folder path
- the `split_report.txt` path
- the generated file count
- any unused names or unwritten chunks shown by the CLI

After a successful simple merge run, report:

- the merged PDF path
- the output folder path
- the `merge_report.txt` path
- the total merged pages

After a successful batch merge run, report:

- the output folder path
- the `merge_report.txt` path
- the generated file count
- a few example output filenames if shown by the CLI

If example output filenames are printed in split mode, you may relay a few of them to confirm naming.

## Do not do these things

- Do not bypass the interactive prompt flow.
- Do not invent nonexistent flags, APIs, or library entrypoints.
- Do not claim MCP support exists in this project.
- Do not use `.command` launchers for normal AI-driven CLI operation.
- Do not auto-confirm warnings without user visibility.
- Do not skip the validation step in automation when `--dry-run --json` is viable.
- Do not hardcode or expose user-specific local paths in repo-tracked skill content.
