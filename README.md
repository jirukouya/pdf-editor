# PDF EDITOR

PDF EDITOR is a macOS-focused PDF splitting tool built around an interactive CLI workflow.

It takes a merged PDF plus a CSV or XLSX file, matches rows in order, and exports smaller PDFs named from each person in the sheet.

## Current Status

Implemented now:

- Interactive CLI workflow
- macOS double-click launchers
- CSV input support
- XLSX input support
- automatic name column detection
- optional order column detection
- filename suffix input
- automatic output folder creation
- duplicate filename auto-renaming
- text report generation

Not packaged yet inside this repository:

- a standalone MCP server
- built-in Agent Skills files
- native `.app` packaging

That means the project is currently CLI-first. MCP and Agent Skills are best treated as integration layers around this CLI.

## Recommended Use on macOS

For non-technical users:

1. Double-click `Setup PDF Editor.command` once
2. Wait for setup to finish
3. Double-click `Launch PDF Editor.command`

This opens Terminal automatically and runs the tool for you.

## First-Time Setup

`Setup PDF Editor.command` will:

- create a local `.venv` virtual environment
- install PDF EDITOR and its required dependencies

After that, users normally only need to open `Launch PDF Editor.command`.

## CLI Workflow

The current CLI asks for:

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

## MCP Integration

This repository does not currently include an MCP server implementation.

If you want to use PDF EDITOR in an MCP-based workflow, the recommended approach is:

1. keep this project as the local execution tool
2. wrap it with a lightweight MCP server later
3. let the MCP server call the CLI or the Python entry point

Recommended MCP responsibility split:

- MCP server: receive structured requests from the client or agent
- PDF EDITOR CLI: perform the actual PDF split job
- report file: return a simple artifact the MCP server can reference

A practical MCP action shape would be something like:

- `split_pdf_named`
- inputs: `sheet_path`, `pdf_path`, `pages_per_file`, `suffix`, `output_dir`
- outputs: generated folder path, written file count, warnings, report path

If you want to build MCP support later, the best next step is to extract the split flow in `pdf_editor/app.py` into a cleaner reusable service layer, then expose that through a separate MCP adapter.

## Agent Skills

This repository does not currently ship with Codex/agent skill files, but it is a good candidate for one.

A useful Agent Skill for this project would tell an agent to:

1. ask the user for the merged PDF path
2. ask the user for the CSV/XLSX path
3. ask for pages per split
4. ask for the filename suffix
5. ask for the output folder, or leave it blank
6. run `python3 -m pdf_editor` or call the underlying Python module directly
7. summarize the result and point to the generated report

Recommended Agent Skill scope:

- use the project only for local PDF splitting jobs
- do not guess file paths when the user has not provided them
- confirm mismatched page counts vs sheet rows before continuing
- surface the generated `split_report.txt` in the final response

In other words, Agent Skills should orchestrate the interaction, while PDF EDITOR remains the execution engine.

## GitHub Use

Recommended structure:

- GitHub repository: store source code, tests, launchers, and docs
- GitHub Release: upload the ready-to-use zip file for end users

Keep these in the repository:

- `pdf_editor/`
- `tests/`
- `Setup PDF Editor.command`
- `Launch PDF Editor.command`
- `Create Release Zip.command`
- `README.md`
- `RELEASE_NOTES_v0.1.0.md`
- `GITHUB_RELEASE_TEMPLATE.md`
- `GITHUB_REPO_DESCRIPTION.txt`
- `pyproject.toml`
- `.gitignore`
- `LICENSE`

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

## Notes

- The project currently targets macOS.
- XLSX support does not require `openpyxl`.
- If users download the project from GitHub, macOS may ask them to confirm opening `.command` files the first time.

## License

This project is released under the MIT License. See `LICENSE` for details.
