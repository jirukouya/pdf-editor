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
# PDF EDITOR v0.1.0

PDF EDITOR is a macOS-focused tool for splitting PDFs from CSV/XLSX naming data and merging two PDFs into one output file.

## Highlights

- Interactive step-by-step CLI workflow
- CSV and XLSX input support
- custom naming template with `{Name}` placeholder
- merge two PDFs into one output file
- automatic output folder creation
- duplicate filename auto-renaming
- text report generation
- macOS double-click setup and launch files
- Python 3.9+ compatibility

## Download and Use

1. Download `PDF Editor.zip` from this release
2. Extract the zip file
3. Double-click `Setup PDF Editor.command`
4. After setup finishes, double-click `Launch PDF Editor.command`

## Notes

- macOS may ask you to confirm opening `.command` files the first time.
- The setup script creates a local virtual environment and installs PDF EDITOR with its dependencies.
- This release is intended for macOS.
```
