from __future__ import annotations

import argparse
import csv
import importlib
import re
import shlex
import subprocess
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile


BANNER = r"""
==================================================

██████╗ ██████╗ ███████╗
██╔══██╗██╔══██╗██╔════╝
██████╔╝██║  ██║█████╗
██╔═══╝ ██║  ██║██╔══╝
██║     ██████╔╝██║
╚═╝     ╚═════╝ ╚═╝

███████╗██████╗ ██╗████████╗ ██████╗ ██████╗
██╔════╝██╔══██╗██║╚══██╔══╝██╔═══██╗██╔══██╗
█████╗  ██║  ██║██║   ██║   ██║   ██║██████╔╝
██╔══╝  ██║  ██║██║   ██║   ██║   ██║██╔══██╗
███████╗██████╔╝██║   ██║   ╚██████╔╝██║  ██║
╚══════╝╚═════╝ ╚═╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝

==================================================
"""

OOXML_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rels": "http://schemas.openxmlformats.org/package/2006/relationships",
}
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SUPPORTED_SHEET_EXTENSIONS = {".csv", ".xlsx"}

NAME_CANDIDATES = [
    "name",
    "employee name",
    "full name",
    "staff name",
]

ORDER_CANDIDATES = [
    "no",
    "number",
    "id",
    "employee id",
    "staff id",
]


@dataclass(slots=True)
class InputRecord:
    index: int
    order: int
    name: str


@dataclass(slots=True)
class JobConfig:
    sheet_path: Path
    pdf_path: Path
    pages_per_file: int
    suffix: str
    output_dir: Path
    name_column: str
    order_column: str | None


@dataclass(slots=True)
class SplitResult:
    written: int
    skipped_names: int
    skipped_chunks: int
    output_files: list[Path]


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="pdf-editor",
        description="Interactive PDF splitting CLI.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="pdf-editor 0.1.0",
    )
    parser.add_argument(
        "--simulate-missing-deps",
        default="",
        help="Developer testing only: comma-separated module names to simulate as missing during the first startup check.",
    )
    args = parser.parse_args()
    simulated_missing = parse_simulated_missing_deps(args.simulate_missing_deps)
    try:
        run_interactive(simulated_missing)
    except KeyboardInterrupt:
        print("\nCancelled.")
        raise SystemExit(130)


def run_interactive(simulated_missing: list[str] | None = None) -> None:
    print(BANNER)
    print("Welcome to PDF EDITOR.")
    print("I will guide you step by step and split your PDF for you.\n")
    run_startup_checks(simulated_missing)

    sheet_path = prompt_existing_file(
        "[1/5] Where is your CSV/XLSX file?",
        allowed_extensions=SUPPORTED_SHEET_EXTENSIONS,
    )
    fieldnames = inspect_sheet(sheet_path)
    name_column = pick_column(fieldnames, NAME_CANDIDATES)
    order_column = pick_column(fieldnames, ORDER_CANDIDATES)

    print(f"Loaded sheet: {sheet_path.name}")
    print("Detected columns:")
    for column in fieldnames:
        print(f"- {column}")
    print("")

    if not name_column:
        print("I could not automatically detect the name column.")
        name_column = prompt_column_choice(
            "Please enter the correct name column:",
            fieldnames,
        )

    print(f"Name column: {name_column}")
    if order_column:
        print(f"Order column: {order_column}")
    else:
        print("No order column detected. I will use the original row order.")
    print("")

    if not prompt_yes_no("Are these column settings correct?", default=True):
        name_column = prompt_column_choice(
            "Please enter the correct name column:",
            fieldnames,
        )
        order_column = prompt_optional_column_choice(
            "If you want a specific order column, enter it now. Otherwise press Enter:",
            fieldnames,
        )

    _, records, _, _ = read_sheet_records(
        sheet_path,
        forced_name_column=name_column,
        forced_order_column=order_column,
    )

    pdf_path = prompt_existing_file(
        "\n[2/5] Where is your PDF file?",
        allowed_extensions={".pdf"},
    )
    total_pages = get_pdf_page_count(pdf_path)
    print(f"Loaded PDF. Total pages: {total_pages}")

    pages_per_file = prompt_positive_int(
        "\n[3/5] How many pages should each split PDF contain?",
        default=1,
    )
    print(f"Each output PDF will contain {pages_per_file} page(s).")

    suffix = input(
        "\n[4/5] Enter the filename suffix after the person's name. Leave blank if not needed.\n"
        "Example input: EA Form Revised\n"
        "Example output: Alice Tan - EA Form Revised.pdf\n"
        "If left blank: Alice Tan.pdf\n"
        "> "
    ).strip()
    suffix = sanitize_suffix(suffix)

    example_names = [record.name for record in records[:3]]
    if example_names:
        print("\nFilename preview:")
        for name in example_names:
            print(f"- {build_output_filename(name, suffix)}")

    output_dir = prompt_output_dir(
        "\n[5/5] Where should I save the generated PDFs?\n"
        "Leave blank and I will create an output folder automatically based on your filename suffix.\n"
        "> ",
        pdf_path,
        suffix,
    )

    config = JobConfig(
        sheet_path=sheet_path,
        pdf_path=pdf_path,
        pages_per_file=pages_per_file,
        suffix=suffix,
        output_dir=output_dir,
        name_column=name_column,
        order_column=order_column,
    )

    warnings = build_warnings(records, total_pages, pages_per_file)
    show_summary(config, total_pages, len(records), warnings)

    if not prompt_yes_no("\nDo you want to start generating PDFs now?", default=True):
        print("Cancelled.")
        return

    result = split_pdf_named(config, records, total_pages)
    write_report(config, total_pages, len(records), warnings, result)
    show_completion(config, result)


def run_startup_checks(simulated_missing: list[str] | None = None) -> None:
    missing = find_missing_dependencies(simulated_missing=simulated_missing)
    if missing:
        print("Startup check found a missing required library:")
        for module_name in missing:
            print(f"- {module_name}")

        if prompt_yes_no("\nDo you want me to install it now?", default=True):
            if not install_missing_dependencies(missing):
                print("\nAutomatic installation failed.")
                print("Please install it manually with:")
                print(f"{sys.executable} -m pip install {' '.join(missing)}")
                raise SystemExit(1)

            missing = find_missing_dependencies()
            if missing:
                print("\nInstallation finished, but these libraries are still missing:")
                for module_name in missing:
                    print(f"- {module_name}")
                print("Please install them manually and try again.")
                raise SystemExit(1)

            print("\nInstallation completed successfully.")
        else:
            print("\nPlease install it manually with:")
            print(f"{sys.executable} -m pip install {' '.join(missing)}")
            raise SystemExit(1)

    print("Startup check passed. Required libraries are installed.\n")


def find_missing_dependencies(
    module_loader: Callable[[str], object] | None = None,
    simulated_missing: list[str] | None = None,
) -> list[str]:
    loader = module_loader or importlib.import_module
    required_modules = ["pypdf"]
    missing: list[str] = []
    simulated_missing_set = {name.strip() for name in (simulated_missing or []) if name.strip()}

    for module_name in required_modules:
        if module_name in simulated_missing_set:
            missing.append(module_name)
            continue
        try:
            loader(module_name)
        except ModuleNotFoundError:
            missing.append(module_name)

    return missing


def parse_simulated_missing_deps(raw: str) -> list[str]:
    return [value.strip() for value in raw.split(",") if value.strip()]


def install_missing_dependencies(
    module_names: list[str],
    installer: Callable[[list[str]], int] | None = None,
) -> bool:
    run_installer = installer or run_dependency_installer
    return run_installer(module_names) == 0


def run_dependency_installer(module_names: list[str]) -> int:
    command = [sys.executable, "-m", "pip", "install", *module_names]
    try:
        completed = subprocess.run(command, check=False)
    except OSError:
        return 1
    return completed.returncode


def prompt_existing_file(message: str, allowed_extensions: set[str] | None = None) -> Path:
    while True:
        raw = input(f"{message}\n> ")
        path = parse_path_input(raw)
        if not path:
            print("Please enter a valid path.")
            continue
        if not path.exists():
            print("That file was not found. Please try again.")
            continue
        if not path.is_file():
            print("That path is not a file. Please try again.")
            continue
        if allowed_extensions and path.suffix.casefold() not in allowed_extensions:
            allowed = ", ".join(sorted(allowed_extensions))
            print(f"Please provide a supported file type: {allowed}")
            continue
        return path


def prompt_positive_int(message: str, default: int) -> int:
    while True:
        raw = input(f"{message}\nDefault is {default}. Press Enter to use it.\n> ").strip()
        if not raw:
            return default
        try:
            value = int(raw)
        except ValueError:
            print("Please enter a whole number.")
            continue
        if value <= 0:
            print("Please enter a number greater than 0.")
            continue
        return value


def prompt_output_dir(message: str, pdf_path: Path, suffix: str) -> Path:
    raw = input(message)
    path = parse_path_input(raw)
    if path:
        return path
    auto_dir = build_default_output_dir(pdf_path, suffix)
    print("\nNo output folder was provided.")
    if suffix:
        print(f'Default folder name will follow your filename suffix: "{suffix}"')
    else:
        print(f'Default folder name will follow the source PDF name: "{pdf_path.stem}"')
    print(f"I will create this folder automatically:\n{auto_dir}")
    return auto_dir


def prompt_yes_no(message: str, default: bool) -> bool:
    suffix = "(Y/n)" if default else "(y/N)"
    while True:
        raw = input(f"{message} {suffix}\n> ").strip().lower()
        if not raw:
            return default
        if raw in {"y", "yes"}:
            return True
        if raw in {"n", "no"}:
            return False
        print("Please enter y or n.")


def prompt_column_choice(message: str, fieldnames: list[str]) -> str:
    available = {name.casefold(): name for name in fieldnames}
    while True:
        raw = input(f"{message}\n> ").strip()
        if raw.casefold() in available:
            return available[raw.casefold()]
        print("That column was not found. Please enter one of the detected column names.")


def prompt_optional_column_choice(message: str, fieldnames: list[str]) -> str | None:
    available = {name.casefold(): name for name in fieldnames}
    while True:
        raw = input(f"{message}\n> ").strip()
        if not raw:
            return None
        if raw.casefold() in available:
            return available[raw.casefold()]
        print("That column was not found. Please try again or press Enter to skip.")


def parse_path_input(raw: str) -> Path | None:
    raw = raw.strip()
    if not raw:
        return None
    try:
        parts = shlex.split(raw)
    except ValueError:
        parts = [raw.strip('"').strip("'")]
    value = parts[0] if parts else raw
    return Path(value).expanduser()


def inspect_sheet(sheet_path: Path) -> list[str]:
    fieldnames, _ = load_sheet_rows(sheet_path)
    return fieldnames


def read_sheet_records(
    sheet_path: Path,
    forced_name_column: str | None = None,
    forced_order_column: str | None = None,
) -> tuple[list[str], list[InputRecord], str, str | None]:
    fieldnames, rows = load_sheet_rows(sheet_path)
    if not fieldnames:
        raise SystemExit("The sheet does not contain a header row.")

    name_column = forced_name_column or pick_column(fieldnames, NAME_CANDIDATES)
    if not name_column:
        raise SystemExit(
            "I could not find a name column. Please check for a column like NAME or Full Name."
        )

    order_column = forced_order_column
    if forced_order_column is None:
        order_column = pick_column(fieldnames, ORDER_CANDIDATES)

    records: list[InputRecord] = []
    for index, row in enumerate(rows, start=1):
        name = (row.get(name_column) or "").strip()
        if not name:
            continue
        order = index
        if order_column:
            raw_order = (row.get(order_column) or "").strip()
            try:
                order = int(raw_order)
            except ValueError:
                order = 10**9 + index
        records.append(InputRecord(index=index, order=order, name=name))

    records.sort(key=lambda record: (record.order, record.index))
    return fieldnames, records, name_column, order_column


def load_sheet_rows(sheet_path: Path) -> tuple[list[str], list[dict[str, str]]]:
    suffix = sheet_path.suffix.casefold()
    if suffix == ".csv":
        return load_csv_rows(sheet_path)
    if suffix == ".xlsx":
        return load_xlsx_rows(sheet_path)
    raise SystemExit("Unsupported sheet format. Please use CSV or XLSX.")


def load_csv_rows(sheet_path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with sheet_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fieldnames = reader.fieldnames or []
        if not fieldnames:
            raise SystemExit("The CSV file does not contain a header row.")
        rows = [{key: (value or "") for key, value in row.items()} for row in reader]
    return fieldnames, rows


def load_xlsx_rows(sheet_path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with ZipFile(sheet_path) as workbook:
        shared_strings = read_shared_strings(workbook)
        sheet_xml_path = get_first_sheet_xml_path(workbook)
        root = ET.fromstring(workbook.read(sheet_xml_path))

    rows = root.findall(".//main:sheetData/main:row", OOXML_NS)
    if not rows:
        raise SystemExit("The XLSX file does not contain any rows.")

    table_rows: list[list[str]] = [read_xlsx_row(row, shared_strings) for row in rows]
    header_values = trim_trailing_blanks(table_rows[0])
    fieldnames = [
        value.strip() if value.strip() else f"Column {index + 1}"
        for index, value in enumerate(header_values)
    ]
    if not fieldnames:
        raise SystemExit("The XLSX file does not contain a usable header row.")

    data_rows: list[dict[str, str]] = []
    for row_values in table_rows[1:]:
        trimmed_values = trim_trailing_blanks(row_values)
        row_dict = {
            fieldnames[index]: (trimmed_values[index].strip() if index < len(trimmed_values) else "")
            for index in range(len(fieldnames))
        }
        data_rows.append(row_dict)

    return fieldnames, data_rows


def read_shared_strings(workbook: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []
    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for item in root.findall("main:si", OOXML_NS):
        text = "".join(item.itertext())
        values.append(text)
    return values


def get_first_sheet_xml_path(workbook: ZipFile) -> str:
    workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
    rels_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))

    rel_map = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall("rels:Relationship", OOXML_NS)
    }
    first_sheet = workbook_root.find("main:sheets/main:sheet", OOXML_NS)
    if first_sheet is None:
        raise SystemExit("The XLSX file does not contain any worksheets.")

    rel_id = first_sheet.attrib.get(f"{{{OFFICE_REL_NS}}}id")
    target = rel_map.get(rel_id or "")
    if not target:
        raise SystemExit("The first worksheet could not be resolved from the XLSX file.")
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return f"xl/{target}"


def read_xlsx_row(row: ET.Element, shared_strings: list[str]) -> list[str]:
    values_by_index: dict[int, str] = {}
    for cell in row.findall("main:c", OOXML_NS):
        reference = cell.attrib.get("r", "")
        column_index = column_letters_to_index(reference)
        values_by_index[column_index] = read_xlsx_cell_value(cell, shared_strings)

    if not values_by_index:
        return []
    max_index = max(values_by_index)
    return [values_by_index.get(index, "") for index in range(max_index + 1)]


def read_xlsx_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        inline_root = cell.find("main:is", OOXML_NS)
        return "" if inline_root is None else "".join(inline_root.itertext())

    raw_value = cell.findtext("main:v", default="", namespaces=OOXML_NS)
    if cell_type == "s":
        if raw_value.isdigit():
            shared_index = int(raw_value)
            if 0 <= shared_index < len(shared_strings):
                return shared_strings[shared_index]
        return ""
    return raw_value or ""


def trim_trailing_blanks(values: list[str]) -> list[str]:
    trimmed = list(values)
    while trimmed and not trimmed[-1].strip():
        trimmed.pop()
    return trimmed


def column_letters_to_index(cell_reference: str) -> int:
    match = re.match(r"([A-Z]+)", cell_reference.upper())
    if not match:
        return 0
    letters = match.group(1)
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def pick_column(fieldnames: Iterable[str], candidates: Iterable[str]) -> str | None:
    normalized = {normalize_key(name): name for name in fieldnames}
    for candidate in candidates:
        match = normalized.get(normalize_key(candidate))
        if match:
            return match
    return None


def normalize_key(value: str) -> str:
    return re.sub(r"[\s_\-]+", "", value).casefold()


def get_pdf_page_count(pdf_path: Path) -> int:
    pdf_reader, _ = load_pdf_tools()
    reader = pdf_reader(str(pdf_path))
    return len(reader.pages)


def build_output_filename(name: str, suffix: str) -> str:
    safe_name = sanitize_filename(name)
    safe_suffix = sanitize_suffix(suffix)
    if safe_suffix:
        return f"{safe_name} - {safe_suffix}.pdf"
    return f"{safe_name}.pdf"


def sanitize_filename(name: str) -> str:
    value = re.sub(r'[\\/:*?"<>|]+', " ", name).strip()
    value = re.sub(r"\s+", " ", value)
    return value or "Unknown"


def sanitize_suffix(suffix: str) -> str:
    value = re.sub(r'[\\/:*?"<>|]+', " ", suffix).strip()
    value = re.sub(r"\s+", " ", value)
    if value.casefold().endswith(".pdf"):
        value = value[:-4].rstrip()
    return value


def build_warnings(records: list[InputRecord], total_pages: int, pages_per_file: int) -> list[str]:
    warnings: list[str] = []
    chunk_count = (total_pages + pages_per_file - 1) // pages_per_file
    if len(records) != chunk_count:
        warnings.append(
            f"Sheet record count ({len(records)}) does not match PDF chunk count ({chunk_count})."
        )

    duplicate_names = [
        name for name, count in Counter(record.name for record in records).items() if count > 1
    ]
    if duplicate_names:
        preview = ", ".join(sorted(duplicate_names)[:5])
        if len(duplicate_names) > 5:
            preview += " ..."
        warnings.append(
            "Duplicate names were found. Output files will be auto-renamed: " + preview
        )
    return warnings


def show_summary(
    config: JobConfig,
    total_pages: int,
    record_count: int,
    warnings: list[str],
) -> None:
    chunk_count = (total_pages + config.pages_per_file - 1) // config.pages_per_file
    print("\n------------------------------------------------------------")
    print("Review this summary before I start:")
    print("------------------------------------------------------------")
    print(f"PDF total pages         : {total_pages}")
    print(f"Sheet record count      : {record_count}")
    print(f"Pages per output file   : {config.pages_per_file}")
    print(f"Expected output files   : {min(record_count, chunk_count)}")
    print(f"Name column             : {config.name_column}")
    print(
        f"Order column            : {config.order_column if config.order_column else 'Original row order'}"
    )
    print(
        f"Filename suffix         : {config.suffix if config.suffix else '(blank, name only)'}"
    )
    print(f"Output folder           : {config.output_dir}")
    print("\nChecks:")
    if warnings:
        for warning in warnings:
            print(f"- {warning}")
    else:
        print("- No obvious issues found")


def split_pdf_named(config: JobConfig, records: list[InputRecord], total_pages: int) -> SplitResult:
    pdf_reader, pdf_writer = load_pdf_tools()
    reader = pdf_reader(str(config.pdf_path))
    config.output_dir.mkdir(parents=True, exist_ok=True)

    chunk_count = (total_pages + config.pages_per_file - 1) // config.pages_per_file
    limit = min(len(records), chunk_count)
    output_files: list[Path] = []

    for idx in range(limit):
        start = idx * config.pages_per_file
        end = min(start + config.pages_per_file, total_pages)

        writer = pdf_writer()
        for page_number in range(start, end):
            writer.add_page(reader.pages[page_number])

        base_name = build_output_filename(records[idx].name, config.suffix)
        output_path = ensure_unique_path(config.output_dir / base_name)
        with output_path.open("wb") as handle:
            writer.write(handle)
        output_files.append(output_path)
        print_progress(idx + 1, limit)

    print("")
    return SplitResult(
        written=limit,
        skipped_names=max(0, len(records) - limit),
        skipped_chunks=max(0, chunk_count - limit),
        output_files=output_files,
    )


def load_pdf_tools() -> tuple[type, type]:
    module = importlib.import_module("pypdf")
    return module.PdfReader, module.PdfWriter


def ensure_unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def ensure_unique_directory_path(path: Path) -> Path:
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.parent / f"{path.name} ({counter})"
        if not candidate.exists():
            return candidate
        counter += 1


def build_default_output_dir(pdf_path: Path, suffix: str) -> Path:
    folder_name = sanitize_directory_name(suffix) or sanitize_directory_name(pdf_path.stem)
    folder_name = folder_name or "output"
    return ensure_unique_directory_path(pdf_path.with_name(folder_name))


def sanitize_directory_name(name: str) -> str:
    value = re.sub(r'[\\/:*?"<>|]+', " ", name).strip()
    return re.sub(r"\s+", " ", value).rstrip(".")


def print_progress(current: int, total: int) -> None:
    width = 50
    filled = width if total == 0 else int(width * current / total)
    bar = "#" * filled + "-" * (width - filled)
    print(f"\r[{bar}] {current}/{total}", end="", flush=True)


def write_report(
    config: JobConfig,
    total_pages: int,
    record_count: int,
    warnings: list[str],
    result: SplitResult,
) -> None:
    report_path = config.output_dir / "split_report.txt"
    chunk_count = (total_pages + config.pages_per_file - 1) // config.pages_per_file

    lines = [
        "PDF EDITOR REPORT",
        "",
        f"Sheet path: {config.sheet_path}",
        f"PDF path: {config.pdf_path}",
        f"Output dir: {config.output_dir}",
        f"Pages per file: {config.pages_per_file}",
        f"Suffix: {config.suffix}",
        f"Name column: {config.name_column}",
        f"Order column: {config.order_column or 'Original row order'}",
        "",
        f"PDF total pages: {total_pages}",
        f"Sheet record count: {record_count}",
        f"Chunk count: {chunk_count}",
        f"Written files: {result.written}",
        f"Unused names: {result.skipped_names}",
        f"Unwritten chunks: {result.skipped_chunks}",
        "",
        "Warnings:",
    ]
    if warnings:
        lines.extend(f"- {warning}" for warning in warnings)
    else:
        lines.append("- None")

    lines.extend(["", "Output files:"])
    lines.extend(f"- {path.name}" for path in result.output_files)
    report_path.write_text("\n".join(lines), encoding="utf-8")


def show_completion(config: JobConfig, result: SplitResult) -> None:
    print("------------------------------------------------------------")
    print("Done")
    print("------------------------------------------------------------")
    print(f"Generated PDF files     : {result.written}")
    print(f"Unused names            : {result.skipped_names}")
    print(f"Unwritten chunks        : {result.skipped_chunks}")
    print(f"Output folder           : {config.output_dir}")
    print(f"Report file             : {config.output_dir / 'split_report.txt'}")

    if result.output_files:
        print("\nExample files:")
        for path in result.output_files[:3]:
            print(f"- {path.name}")
