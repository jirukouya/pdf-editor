from __future__ import annotations

import argparse
import csv
import importlib
import os
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
PROJECT_ROOT = Path(__file__).resolve().parent.parent
LOCAL_VENV_PYTHON = PROJECT_ROOT / ".venv" / "bin" / "python"

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
NAME_PLACEHOLDER_PATTERN = re.compile(r"\{\s*name\s*\}", re.IGNORECASE)


@dataclass
class InputRecord:
    index: int
    order: int
    name: str


@dataclass
class JobConfig:
    sheet_path: Path
    pdf_path: Path
    pages_per_file: int
    naming_template: str
    output_dir: Path
    name_column: str
    order_column: str | None


@dataclass
class SplitResult:
    written: int
    skipped_names: int
    skipped_chunks: int
    output_files: list[Path]


@dataclass
class MergeConfig:
    first_pdf_path: Path
    second_pdf_path: Path
    output_path: Path


@dataclass
class MergeResult:
    output_path: Path
    total_pages: int


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="pdf-editor",
        description="Interactive PDF split-and-merge CLI.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="pdf-editor 0.2.0",
    )
    parser.add_argument(
        "--simulate-missing-deps",
        default="",
        help="Developer testing only: comma-separated module names to simulate as missing during the first startup check.",
    )
    parser.add_argument(
        "--mode",
        choices=("split", "merge"),
        help="Run without interactive prompts using the provided mode-specific options.",
    )
    parser.add_argument("--sheet-path", help="CSV/XLSX input path for split mode.")
    parser.add_argument("--pdf-path", help="Source PDF path for split mode.")
    parser.add_argument(
        "--pages-per-file",
        type=int,
        default=1,
        help="Pages per output PDF in split mode.",
    )
    parser.add_argument(
        "--naming-template",
        default="{Name}",
        help="Naming template for split mode. Must include {Name}.",
    )
    parser.add_argument("--output-dir", help="Output directory for split mode.")
    parser.add_argument("--name-column", help="Optional explicit name column for split mode.")
    parser.add_argument("--order-column", help="Optional explicit order column for split mode.")
    parser.add_argument("--first-pdf-path", help="First PDF path for merge mode.")
    parser.add_argument("--second-pdf-path", help="Second PDF path for merge mode.")
    parser.add_argument("--output-path", help="Merged PDF output path for merge mode.")
    args = parser.parse_args()
    simulated_missing = parse_simulated_missing_deps(args.simulate_missing_deps)
    try:
        if args.mode:
            run_non_interactive(args, simulated_missing)
        else:
            run_interactive(simulated_missing)
    except KeyboardInterrupt:
        print("\nCancelled.")
        raise SystemExit(130)


def run_interactive(simulated_missing: list[str] | None = None) -> None:
    print(BANNER)
    print("Welcome to PDF EDITOR.")
    print("I can help you split or merge PDF files step by step.\n")
    run_startup_checks(simulated_missing)

    operation = prompt_operation()
    if operation == "merge":
        run_merge_interactive()
        return
    run_split_interactive()


def run_non_interactive(args: argparse.Namespace, simulated_missing: list[str] | None = None) -> None:
    run_startup_checks(simulated_missing, interactive=False)
    if args.mode == "split":
        run_split_non_interactive(args)
        return
    run_merge_non_interactive(args)


def run_split_non_interactive(args: argparse.Namespace) -> None:
    if not args.sheet_path or not args.pdf_path:
        raise SystemExit("Split mode requires --sheet-path and --pdf-path.")
    if args.pages_per_file <= 0:
        raise SystemExit("--pages-per-file must be greater than 0.")

    sheet_path = validate_existing_file_path(
        Path(args.sheet_path).expanduser(),
        SUPPORTED_SHEET_EXTENSIONS,
        "sheet file",
    )
    pdf_path = validate_existing_file_path(
        Path(args.pdf_path).expanduser(),
        {".pdf"},
        "PDF file",
    )
    naming_template = sanitize_naming_template(args.naming_template or "{Name}")
    if not contains_name_placeholder(naming_template):
        raise SystemExit("The naming template must include {Name}.")

    fieldnames = inspect_sheet(sheet_path)
    name_column = resolve_requested_column_name(fieldnames, args.name_column, "name")
    order_column = resolve_requested_column_name(fieldnames, args.order_column, "order")
    _, records, detected_name_column, detected_order_column = read_sheet_records(
        sheet_path,
        forced_name_column=name_column,
        forced_order_column=order_column,
    )
    total_pages = get_pdf_page_count(pdf_path)
    output_dir = (
        Path(args.output_dir).expanduser()
        if args.output_dir
        else build_default_output_dir(pdf_path, naming_template)
    )
    config = JobConfig(
        sheet_path=sheet_path,
        pdf_path=pdf_path,
        pages_per_file=args.pages_per_file,
        naming_template=naming_template,
        output_dir=output_dir,
        name_column=detected_name_column,
        order_column=detected_order_column,
    )
    warnings = build_warnings(records, total_pages, args.pages_per_file)
    show_summary(config, total_pages, len(records), warnings)
    result = split_pdf_named(config, records, total_pages)
    write_report(config, total_pages, len(records), warnings, result)
    show_completion(config, result)


def run_merge_non_interactive(args: argparse.Namespace) -> None:
    if not args.first_pdf_path or not args.second_pdf_path:
        raise SystemExit("Merge mode requires --first-pdf-path and --second-pdf-path.")

    first_pdf_path = validate_existing_file_path(
        Path(args.first_pdf_path).expanduser(),
        {".pdf"},
        "first PDF file",
    )
    second_pdf_path = validate_existing_file_path(
        Path(args.second_pdf_path).expanduser(),
        {".pdf"},
        "second PDF file",
    )
    output_path = (
        build_default_merge_output_path(first_pdf_path)
        if not args.output_path
        else normalize_merge_output_path(Path(args.output_path).expanduser(), build_merge_output_filename(first_pdf_path))
    )

    first_total_pages = get_pdf_page_count(first_pdf_path)
    second_total_pages = get_pdf_page_count(second_pdf_path)
    config = MergeConfig(
        first_pdf_path=first_pdf_path,
        second_pdf_path=second_pdf_path,
        output_path=output_path,
    )
    show_merge_summary(config, first_total_pages, second_total_pages)
    result = merge_pdf_files(config)
    write_merge_report(config, first_total_pages, second_total_pages, result)
    show_merge_completion(result)


def run_split_interactive() -> None:
    print("Selected function: Split PDF\n")

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

    naming_template = prompt_naming_template()

    example_names = [record.name for record in records[:3]]
    if example_names:
        print("\nFilename preview:")
        for name in example_names:
            print(f"- {build_output_filename(name, naming_template)}")

    output_dir = prompt_output_dir(
        "\n[5/5] Where should I save the generated PDFs?\n"
        "Leave blank and I will create an output folder automatically based on your naming template.\n"
        "> ",
        pdf_path,
        naming_template,
    )

    config = JobConfig(
        sheet_path=sheet_path,
        pdf_path=pdf_path,
        pages_per_file=pages_per_file,
        naming_template=naming_template,
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


def run_merge_interactive() -> None:
    print("Selected function: Merge PDF\n")

    first_pdf_path = prompt_existing_file(
        "[1/3] Where is your first PDF file?",
        allowed_extensions={".pdf"},
    )
    first_total_pages = get_pdf_page_count(first_pdf_path)
    print(f"Loaded first PDF. Total pages: {first_total_pages}")

    second_pdf_path = prompt_existing_file(
        "\n[2/3] Where is your second PDF file?",
        allowed_extensions={".pdf"},
    )
    second_total_pages = get_pdf_page_count(second_pdf_path)
    print(f"Loaded second PDF. Total pages: {second_total_pages}")

    output_path = prompt_merge_output_path(
        "\n[3/3] Where should I save the merged PDF?\n"
        "Leave blank and I will create a 'Merged PDF' folder automatically.\n"
        "If you enter a folder path, I will use the first PDF filename.\n"
        "> ",
        first_pdf_path,
    )

    config = MergeConfig(
        first_pdf_path=first_pdf_path,
        second_pdf_path=second_pdf_path,
        output_path=output_path,
    )
    show_merge_summary(config, first_total_pages, second_total_pages)

    if not prompt_yes_no("\nDo you want to start merging now?", default=True):
        print("Cancelled.")
        return

    result = merge_pdf_files(config)
    write_merge_report(config, first_total_pages, second_total_pages, result)
    show_merge_completion(result)


def run_startup_checks(
    simulated_missing: list[str] | None = None,
    interactive: bool = True,
) -> None:
    missing = find_missing_dependencies(simulated_missing=simulated_missing)
    if missing:
        print("Startup check found a missing required library:")
        for module_name in missing:
            print(f"- {module_name}")

        if not interactive:
            print("\nPlease run 'Setup PDF Editor.command' or install manually with:")
            print(f"{sys.executable} -m pip install {' '.join(missing)}")
            raise SystemExit(1)

        if prompt_yes_no("\nDo you want me to install it now?", default=True):
            if install_missing_dependencies(missing):
                remaining = find_missing_dependencies()
                if not remaining:
                    print("\nInstallation completed successfully.")
                    print("Startup check passed. Required libraries are installed.\n")
                    return
                missing = remaining

            if setup_local_project_environment():
                print("\nLocal project environment is ready.")
                print("Restarting PDF EDITOR using the local virtual environment...\n")
                restart_with_local_venv()

            print("\nAutomatic installation failed.")
            print("Please run 'Setup PDF Editor.command' or install manually with:")
            print(f"{sys.executable} -m pip install {' '.join(missing)}")
            raise SystemExit(1)

        print("\nPlease run 'Setup PDF Editor.command' or install manually with:")
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


def setup_local_project_environment() -> bool:
    if is_running_inside_local_venv():
        return False

    try:
        create_venv = subprocess.run(
            [sys.executable, "-m", "venv", str(PROJECT_ROOT / ".venv")],
            check=False,
        )
    except OSError:
        return False

    if create_venv.returncode != 0 or not LOCAL_VENV_PYTHON.exists():
        return False

    subprocess.run([str(LOCAL_VENV_PYTHON), "-m", "pip", "install", "--upgrade", "pip"], check=False)
    install_project = subprocess.run(
        [str(LOCAL_VENV_PYTHON), "-m", "pip", "install", "-e", str(PROJECT_ROOT)],
        check=False,
    )
    return install_project.returncode == 0


def is_running_inside_local_venv() -> bool:
    try:
        return Path(sys.executable).resolve() == LOCAL_VENV_PYTHON.resolve()
    except OSError:
        return False


def restart_with_local_venv() -> None:
    if not LOCAL_VENV_PYTHON.exists():
        raise SystemExit(1)
    restart_args = strip_simulated_missing_args(sys.argv[1:])
    os.execv(str(LOCAL_VENV_PYTHON), [str(LOCAL_VENV_PYTHON), "-m", "pdf_editor", *restart_args])


def strip_simulated_missing_args(args: list[str]) -> list[str]:
    cleaned: list[str] = []
    skip_next = False
    for arg in args:
        if skip_next:
            skip_next = False
            continue
        if arg == "--simulate-missing-deps":
            skip_next = True
            continue
        if arg.startswith("--simulate-missing-deps="):
            continue
        cleaned.append(arg)
    return cleaned


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


def prompt_operation() -> str:
    while True:
        raw = input(
            "Which function do you want to use?\n"
            "1. Split PDF\n"
            "2. Merge PDF\n"
            "> "
        ).strip().lower()
        if raw in {"1", "split", "split pdf"}:
            return "split"
        if raw in {"2", "merge", "merge pdf"}:
            return "merge"
        print("Please choose 1 for Split PDF or 2 for Merge PDF.")


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


def prompt_naming_template() -> str:
    default_template = "{Name}"
    while True:
        raw = input(
            "\n[4/5] Enter the full naming template.\n"
            "It must include {Name} so I know where to place each person's name.\n"
            "Press Enter to use the default: {Name}\n"
            "Example: GD Pink Form - Letter of Offer ({Name}) 26-3-2026\n"
            "> "
        ).strip()
        template = sanitize_naming_template(raw or default_template)
        if contains_name_placeholder(template):
            return template
        print("Your naming template must include {Name}.")


def prompt_output_dir(message: str, pdf_path: Path, naming_template: str) -> Path:
    raw = input(message)
    path = parse_path_input(raw)
    if path:
        return path
    auto_dir = build_default_output_dir(pdf_path, naming_template)
    default_folder_label = sanitize_directory_name(build_default_output_dir_label(naming_template))
    print("\nNo output folder was provided.")
    if not default_folder_label:
        print(f'Default folder name will follow the source PDF name: "{pdf_path.stem}"')
    else:
        print(f'Default folder name will follow your naming template: "{default_folder_label}"')
    print(f"I will create this folder automatically:\n{auto_dir}")
    return auto_dir


def prompt_merge_output_path(message: str, first_pdf_path: Path) -> Path:
    default_filename = build_merge_output_filename(first_pdf_path)
    raw = input(message)
    path = parse_path_input(raw)
    if not path:
        output_path = build_default_merge_output_path(first_pdf_path)
        print("\nNo output path was provided.")
        print('Default folder name will be: "Merged PDF"')
        print(f'Default file name will follow the first PDF: "{default_filename}"')
        print(f"I will create this file automatically:\n{output_path}")
        return output_path

    output_path = normalize_merge_output_path(path, default_filename)
    if path.exists() and path.is_dir():
        print(f'I will save the merged PDF in that folder using: "{output_path.name}"')
    return output_path


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


def validate_existing_file_path(
    path: Path,
    allowed_extensions: set[str] | None = None,
    label: str = "file",
) -> Path:
    if not path.exists():
        raise SystemExit(f"The {label} was not found: {path}")
    if not path.is_file():
        raise SystemExit(f"The {label} is not a file: {path}")
    if allowed_extensions and path.suffix.casefold() not in allowed_extensions:
        allowed = ", ".join(sorted(allowed_extensions))
        raise SystemExit(f"The {label} must use one of these extensions: {allowed}")
    return path


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


def resolve_requested_column_name(
    fieldnames: list[str],
    requested: str | None,
    label: str,
) -> str | None:
    if not requested:
        return None
    match = pick_column(fieldnames, [requested])
    if match:
        return match
    raise SystemExit(f'The requested {label} column was not found: "{requested}"')


def normalize_key(value: str) -> str:
    return re.sub(r"[\s_\-]+", "", value).casefold()


def get_pdf_page_count(pdf_path: Path) -> int:
    pdf_reader, _ = load_pdf_tools()
    reader = pdf_reader(str(pdf_path))
    return len(reader.pages)


def build_output_filename(name: str, naming_template: str) -> str:
    rendered = render_naming_template(name, naming_template)
    return f"{sanitize_filename(rendered)}.pdf"


def sanitize_filename(name: str) -> str:
    value = re.sub(r'[\\/:*?"<>|]+', " ", name).strip()
    value = re.sub(r"\s+", " ", value)
    return value or "Unknown"


def sanitize_naming_template(naming_template: str) -> str:
    value = naming_template.strip()
    value = re.sub(r"\s+", " ", value)
    if value.casefold().endswith(".pdf"):
        value = value[:-4].rstrip()
    return value


def contains_name_placeholder(naming_template: str) -> bool:
    return bool(NAME_PLACEHOLDER_PATTERN.search(naming_template))


def render_naming_template(name: str, naming_template: str) -> str:
    safe_name = sanitize_filename(name)
    template = sanitize_naming_template(naming_template)
    if not contains_name_placeholder(template):
        return safe_name
    rendered = NAME_PLACEHOLDER_PATTERN.sub(safe_name, template)
    return re.sub(r"\s+", " ", rendered).strip() or safe_name


def build_default_output_dir_label(naming_template: str) -> str:
    template = sanitize_naming_template(naming_template)
    value = NAME_PLACEHOLDER_PATTERN.sub("", template)
    value = re.sub(r"\(\s*\)", "", value)
    value = re.sub(r"\[\s*\]", "", value)
    value = re.sub(r"\s+", " ", value).strip(" -_")
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
        f"Naming template         : {config.naming_template}"
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

        base_name = build_output_filename(records[idx].name, config.naming_template)
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


def show_merge_summary(config: MergeConfig, first_total_pages: int, second_total_pages: int) -> None:
    print("\n------------------------------------------------------------")
    print("Review this summary before I start:")
    print("------------------------------------------------------------")
    print(f"First PDF               : {config.first_pdf_path}")
    print(f"First PDF total pages   : {first_total_pages}")
    print(f"Second PDF              : {config.second_pdf_path}")
    print(f"Second PDF total pages  : {second_total_pages}")
    print(f"Merged output file      : {config.output_path}")
    print(f"Total merged pages      : {first_total_pages + second_total_pages}")


def merge_pdf_files(config: MergeConfig) -> MergeResult:
    pdf_reader, pdf_writer = load_pdf_tools()
    first_reader = pdf_reader(str(config.first_pdf_path))
    second_reader = pdf_reader(str(config.second_pdf_path))
    writer = pdf_writer()

    for reader in (first_reader, second_reader):
        for page in reader.pages:
            writer.add_page(page)

    config.output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path = ensure_unique_path(config.output_path)
    with output_path.open("wb") as handle:
        writer.write(handle)

    total_pages = len(first_reader.pages) + len(second_reader.pages)
    return MergeResult(output_path=output_path, total_pages=total_pages)


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


def build_default_output_dir(pdf_path: Path, naming_template: str) -> Path:
    folder_name = sanitize_directory_name(build_default_output_dir_label(naming_template))
    folder_name = folder_name or sanitize_directory_name(pdf_path.stem)
    folder_name = folder_name or "output"
    return ensure_unique_directory_path(pdf_path.with_name(folder_name))


def build_default_merge_output_dir(first_pdf_path: Path) -> Path:
    return ensure_unique_directory_path(first_pdf_path.with_name("Merged PDF"))


def build_merge_output_filename(first_pdf_path: Path) -> str:
    base_name = sanitize_directory_name(first_pdf_path.stem) or "merged"
    return f"{base_name}.pdf"


def build_default_merge_output_path(first_pdf_path: Path) -> Path:
    return ensure_unique_path(
        build_default_merge_output_dir(first_pdf_path) / build_merge_output_filename(first_pdf_path)
    )


def normalize_merge_output_path(path: Path, default_filename: str) -> Path:
    if path.exists() and path.is_dir():
        return ensure_unique_path(path / default_filename)
    if path.suffix.casefold() != ".pdf":
        path = path.with_suffix(".pdf")
    return ensure_unique_path(path)


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
        f"Naming template: {config.naming_template}",
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


def write_merge_report(
    config: MergeConfig,
    first_total_pages: int,
    second_total_pages: int,
    result: MergeResult,
) -> None:
    report_path = result.output_path.parent / "merge_report.txt"
    lines = [
        "PDF EDITOR MERGE REPORT",
        "",
        f"First PDF path: {config.first_pdf_path}",
        f"Second PDF path: {config.second_pdf_path}",
        f"Output file: {result.output_path}",
        "",
        f"First PDF total pages: {first_total_pages}",
        f"Second PDF total pages: {second_total_pages}",
        f"Total merged pages: {result.total_pages}",
    ]
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


def show_merge_completion(result: MergeResult) -> None:
    print("------------------------------------------------------------")
    print("Done")
    print("------------------------------------------------------------")
    print(f"Merged PDF file         : {result.output_path}")
    print(f"Total merged pages      : {result.total_pages}")
    print(f"Output folder           : {result.output_path.parent}")
    print(f"Report file             : {result.output_path.parent / 'merge_report.txt'}")
