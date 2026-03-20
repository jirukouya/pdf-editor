"""MCP server for pdf-editor.

Exposes the PDF splitting functionality as a single MCP tool so that
Claude Code and other MCP-compatible agents can call it without any
interactive prompts.

Usage
-----
Install with the optional dependency group::

    pip install -e ".[mcp]"

Then register in your project's .mcp.json::

    {
      "mcpServers": {
        "pdf-editor": {
          "command": ".venv/bin/pdf-editor-mcp"
        }
      }
    }

Claude can then call the ``split_pdf`` tool directly.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

try:
    from mcp.server import Server
    from mcp.server.stdio import stdio_server
    from mcp import types as mcp_types
except ModuleNotFoundError:
    print(
        "ERROR: The 'mcp' package is not installed.\n"
        f"Run: {sys.executable} -m pip install 'pdf-editor-cli[mcp]'",
        file=sys.stderr,
    )
    raise SystemExit(1)

from pdf_editor.app import (
    SUPPORTED_SHEET_EXTENSIONS,
    JobConfig,
    build_default_output_dir,
    build_warnings,
    find_missing_dependencies,
    get_pdf_page_count,
    read_sheet_records,
    sanitize_suffix,
    split_pdf_named,
    write_report,
)

server: Server = Server("pdf-editor")

SPLIT_PDF_DESCRIPTION = """\
Split a merged PDF into individually-named PDFs using a CSV or XLSX name list.

Required parameters
-------------------
sheet : str
    Absolute or ~ path to the CSV or XLSX file containing the name list.
pdf : str
    Absolute or ~ path to the merged PDF to split.
pages : int
    Number of pages per output file (e.g. 2 for 2-page PDFs).

Optional parameters
-------------------
suffix : str
    Text appended after each person's name, e.g. "EA Form 2024".
    Output becomes "Alice Tan - EA Form 2024.pdf".
output_dir : str
    Directory where split PDFs are saved.  Auto-generated if omitted.
name_column : str
    Exact column header for names.  Auto-detected if omitted.
order_column : str
    Exact column header for ordering.  Auto-detected if omitted.

Returns
-------
JSON with keys: status, written, skipped_names, skipped_chunks,
output_dir, report, warnings, output_files.
"""


@server.list_tools()
async def list_tools() -> list[mcp_types.Tool]:
    return [
        mcp_types.Tool(
            name="split_pdf",
            description=SPLIT_PDF_DESCRIPTION,
            inputSchema={
                "type": "object",
                "required": ["sheet", "pdf", "pages"],
                "properties": {
                    "sheet": {"type": "string", "description": "Path to CSV or XLSX name list."},
                    "pdf": {"type": "string", "description": "Path to the merged PDF."},
                    "pages": {"type": "integer", "minimum": 1, "description": "Pages per output file."},
                    "suffix": {"type": "string", "description": "Optional filename suffix.", "default": ""},
                    "output_dir": {"type": "string", "description": "Output directory (auto if omitted).", "default": ""},
                    "name_column": {"type": "string", "description": "Force a specific name column.", "default": ""},
                    "order_column": {"type": "string", "description": "Force a specific order column.", "default": ""},
                },
            },
        )
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[mcp_types.TextContent]:
    if name != "split_pdf":
        return [mcp_types.TextContent(type="text", text=f"Unknown tool: {name}")]

    errors: list[str] = []

    sheet_str = arguments.get("sheet", "")
    pdf_str = arguments.get("pdf", "")
    pages = arguments.get("pages", 0)

    if not sheet_str:
        errors.append("'sheet' is required.")
    if not pdf_str:
        errors.append("'pdf' is required.")
    if not isinstance(pages, int) or pages <= 0:
        errors.append("'pages' must be a positive integer.")

    if errors:
        return [mcp_types.TextContent(type="text", text="ERROR: " + " ".join(errors))]

    sheet_path = Path(sheet_str).expanduser()
    pdf_path = Path(pdf_str).expanduser()

    for label, path, exts in [
        ("sheet", sheet_path, SUPPORTED_SHEET_EXTENSIONS),
        ("pdf", pdf_path, {".pdf"}),
    ]:
        if not path.exists():
            return [mcp_types.TextContent(type="text", text=f"ERROR: '{label}' path does not exist: {path}")]
        if path.suffix.casefold() not in exts:
            return [mcp_types.TextContent(type="text", text=f"ERROR: '{label}' has unsupported extension '{path.suffix}'.")]

    missing = find_missing_dependencies()
    if missing:
        return [mcp_types.TextContent(type="text", text=f"ERROR: Missing dependencies: {', '.join(missing)}")]

    forced_name = arguments.get("name_column") or None
    forced_order = arguments.get("order_column") or None
    suffix = sanitize_suffix(arguments.get("suffix", ""))
    output_dir_str = arguments.get("output_dir", "")

    try:
        _, records, name_column, order_column = read_sheet_records(
            sheet_path,
            forced_name_column=forced_name,
            forced_order_column=forced_order,
        )
    except SystemExit as exc:
        return [mcp_types.TextContent(type="text", text=f"ERROR reading sheet: {exc}")]

    total_pages = get_pdf_page_count(pdf_path)

    if output_dir_str:
        output_dir = Path(output_dir_str).expanduser()
    else:
        output_dir = build_default_output_dir(pdf_path, suffix)

    config = JobConfig(
        sheet_path=sheet_path,
        pdf_path=pdf_path,
        pages_per_file=pages,
        suffix=suffix,
        output_dir=output_dir,
        name_column=name_column,
        order_column=order_column,
    )

    warnings = build_warnings(records, total_pages, pages)
    result = split_pdf_named(config, records, total_pages)
    write_report(config, total_pages, len(records), warnings, result)

    output = {
        "status": "ok",
        "written": result.written,
        "skipped_names": result.skipped_names,
        "skipped_chunks": result.skipped_chunks,
        "output_dir": str(output_dir),
        "report": str(output_dir / "split_report.txt"),
        "warnings": warnings,
        "output_files": [str(p) for p in result.output_files],
    }
    return [mcp_types.TextContent(type="text", text=json.dumps(output, ensure_ascii=False, indent=2))]


def main() -> None:
    import asyncio

    async def _run() -> None:
        async with stdio_server() as (read_stream, write_stream):
            await server.run(read_stream, write_stream, server.create_initialization_options())

    asyncio.run(_run())


if __name__ == "__main__":
    main()
