# Releasing PDF EDITOR

This file is for maintainers, not end users.

## Repository Description

Suggested GitHub repository description:

Interactive macOS PDF split-and-merge tool with CSV/XLSX naming and double-click launchers.

## Release Package

The end-user release zip should include only:

- `Setup PDF Editor.command`
- `Launch PDF Editor.command`
- `README.md`
- `pdf_editor/`
- `pyproject.toml`

It should not include maintainer-only files such as:

- `tests/`
- `docs/`
- `Create Release Zip.command`
- release templates
- release notes drafts

## Create the Release Zip

Double-click:

- `Create Release Zip.command`

This creates:

- `dist/PDF Editor.zip`

## Suggested Release Text

```md
# PDF EDITOR v0.2.0

PDF EDITOR is a macOS-focused tool for splitting PDFs from CSV/XLSX naming data and merging two PDFs into one output file.

## Highlights

- Split and Merge modes in one CLI
- Custom split naming templates with `{Name}` placeholder
- New Merge PDF workflow for combining two PDFs into one file
- Smart default output handling:
  - Split mode auto-creates output folders
  - Merge mode auto-creates a `Merged PDF` folder when no output path is provided
  - Existing folders and filenames auto-increment with `(2)`, `(3)`, and so on
- AI-friendly fast CLI mode with direct flags for split and merge
- Generated report files for both workflows:
  - `split_report.txt`
  - `merge_report.txt`
- macOS double-click setup and launch files
- Python 3.9+ compatibility

## Download and Use

1. Download `PDF Editor.zip` from this release
2. Extract the zip file
3. Double-click `Setup PDF Editor.command`
4. After setup finishes, double-click `Launch PDF Editor.command`

## Fast CLI For AI And Automation

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

## Notes

- macOS may ask you to confirm opening `.command` files the first time.
- The setup script creates a local virtual environment and installs PDF EDITOR with its dependencies.
- This release is intended for macOS.
```
