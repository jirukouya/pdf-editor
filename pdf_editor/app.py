from __future__ import annotations

import argparse
import csv
import importlib
import json
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
OUTPUT_EXISTS_POLICIES = {"fail", "overwrite", "rename", "continue"}
DUPLICATE_NAME_POLICIES = {"autorename", "fail", "append-row-number", "append-order"}


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
    output_exists_policy: str = "rename"
    duplicate_name_policy: str = "autorename"


@dataclass
class SplitResult:
    written: int
    skipped_names: int
    skipped_chunks: int
    output_files: list[Path]
    overwritten_files: list[Path]
    renamed_files: list[Path]
    skipped_existing_files: list[Path]


@dataclass
class MergeConfig:
    first_pdf_path: Path
    second_pdf_path: Path
    output_path: Path
    merge_order: str = "first-second"
    output_exists_policy: str = "rename"


@dataclass
class MergeResult:
    output_path: Path
    total_pages: int
    overwritten_files: list[Path]
    renamed_files: list[Path]
    skipped_existing_files: list[Path]


@dataclass
class BatchMergeConfig:
    input_dir: Path
    fixed_pdf_path: Path
    merge_order: str
    output_dir: Path
    output_exists_policy: str = "rename"


@dataclass
class BatchMergeResult:
    written: int
    total_pages_per_file: int
    output_files: list[Path]
    overwritten_files: list[Path]
    renamed_files: list[Path]
    skipped_existing_files: list[Path]


@dataclass
class PlannedSplitOutput:
    record: InputRecord
    requested_filename: str
    final_path: Path
    action: str


@dataclass
class PlannedBatchOutput:
    input_pdf_path: Path
    final_path: Path
    action: str


@dataclass
class FastCliPreflight:
    status: str
    phase: str
    mode: str
    merge_kind: str | None
    can_proceed: bool
    requires_confirmation: bool
    warnings: list[str]
    errors: list[str]
    summary: dict[str, object]
    result: dict[str, object] | None = None


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="pdf-editor",
        description="Interactive PDF split-and-merge CLI.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version="pdf-editor 0.2.1",
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
    parser.add_argument(
        "--merge-kind",
        choices=("simple", "batch"),
        default="simple",
        help="Merge workflow to run in merge mode.",
    )
    parser.add_argument("--first-pdf-path", help="First PDF path for merge mode.")
    parser.add_argument("--second-pdf-path", help="Second PDF path for merge mode.")
    parser.add_argument("--output-path", help="Merged PDF output path for merge mode.")
    parser.add_argument("--batch-input-dir", help="Input folder of split PDFs for batch merge mode.")
    parser.add_argument("--fixed-pdf-path", help="Fixed PDF path for batch merge mode.")
    parser.add_argument(
        "--merge-order",
        choices=("split-first", "fixed-first"),
        default="split-first",
        help="Page order for batch merge mode.",
    )
    parser.add_argument("--batch-output-dir", help="Output folder for batch merge mode.")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate inputs and warnings without writing output files.",
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Alias for --dry-run.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit a single JSON object to stdout for fast CLI runs.",
    )
    parser.add_argument(
        "--confirm",
        action="store_true",
        help="Allow execution to proceed when fast CLI preflight returns warnings.",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Treat fast CLI warnings as errors and stop execution.",
    )
    parser.add_argument(
        "--on-output-exists",
        choices=tuple(sorted(OUTPUT_EXISTS_POLICIES)),
        default="fail",
        help="Conflict policy for explicit fast-CLI outputs.",
    )
    parser.add_argument(
        "--duplicate-name-policy",
        choices=tuple(sorted(DUPLICATE_NAME_POLICIES)),
        default="autorename",
        help="Duplicate rendered filename policy for fast split mode.",
    )
    args = parser.parse_args()
    simulated_missing = parse_simulated_missing_deps(args.simulate_missing_deps)
    try:
        if args.mode:
            raise SystemExit(run_non_interactive(args, simulated_missing))
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


def run_non_interactive(args: argparse.Namespace, simulated_missing: list[str] | None = None) -> int:
    args.dry_run = args.dry_run or args.validate_only
    try:
        run_startup_checks(
            simulated_missing,
            interactive=False,
            verbose=not args.json,
        )
        preflight, context = build_fast_cli_preflight(args)
    except SystemExit as exc:
        return emit_fast_cli_error(args, str(exc) or "Unknown error")

    if preflight.status == "error":
        return emit_fast_cli_result(args, preflight, exit_code=1)
    if preflight.status == "warning":
        if args.strict:
            preflight = FastCliPreflight(
                status="error",
                phase="validate",
                mode=preflight.mode,
                merge_kind=preflight.merge_kind,
                can_proceed=False,
                requires_confirmation=False,
                warnings=preflight.warnings,
                errors=["Warnings were treated as errors because --strict was provided."],
                summary=preflight.summary,
            )
            return emit_fast_cli_result(args, preflight, exit_code=1)
        if args.dry_run or not args.confirm:
            return emit_fast_cli_result(args, preflight, exit_code=2)

    if args.dry_run:
        return emit_fast_cli_result(args, preflight, exit_code=0)

    executed = execute_fast_cli_context(context, preflight)
    return emit_fast_cli_result(args, executed, exit_code=0)


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
    merge_kind = prompt_merge_kind()
    if merge_kind == "batch":
        run_batch_merge_interactive()
        return

    print("Selected merge type: Simple Merge\n")

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
        merge_order="first-second",
    )
    show_merge_summary(config, first_total_pages, second_total_pages)

    if not prompt_yes_no("\nDo you want to start merging now?", default=True):
        print("Cancelled.")
        return

    result = merge_pdf_files(config)
    write_merge_report(config, first_total_pages, second_total_pages, result)
    show_merge_completion(result)


def run_batch_merge_interactive() -> None:
    print("Selected merge type: Batch Merge\n")

    input_dir = prompt_existing_directory(
        "[1/4] Where is your split-output folder?",
    )
    input_pdfs = ensure_batch_input_pdfs_exist(input_dir)
    print(f"Found {len(input_pdfs)} PDF file(s) in that folder.")

    fixed_pdf_path = prompt_existing_file(
        "\n[2/4] Where is the fixed PDF file?",
        allowed_extensions={".pdf"},
    )
    fixed_total_pages = get_pdf_page_count(fixed_pdf_path)
    print(f"Loaded fixed PDF. Total pages: {fixed_total_pages}")

    merge_order = prompt_batch_merge_order()
    output_dir = prompt_batch_merge_output_dir(
        "\n[4/4] Where should I save the batch merged PDFs?\n"
        "Leave blank and I will create a 'Batch Merged PDF' folder automatically.\n"
        "> ",
        input_dir,
    )
    config = BatchMergeConfig(
        input_dir=input_dir,
        fixed_pdf_path=fixed_pdf_path,
        merge_order=merge_order,
        output_dir=output_dir,
    )
    show_batch_merge_summary(config)

    if not prompt_yes_no("\nDo you want to start batch merging now?", default=True):
        print("Cancelled.")
        return

    result = merge_pdf_folder(config)
    write_batch_merge_report(config, result)
    show_batch_merge_completion(config, result)


def build_fast_cli_preflight(
    args: argparse.Namespace,
) -> tuple[FastCliPreflight, dict[str, object] | None]:
    if args.mode == "split":
        return build_split_fast_cli_preflight(args)
    if args.merge_kind == "batch":
        return build_batch_merge_fast_cli_preflight(args)
    return build_simple_merge_fast_cli_preflight(args)


def build_split_fast_cli_preflight(
    args: argparse.Namespace,
) -> tuple[FastCliPreflight, dict[str, object] | None]:
    try:
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
        output_exists_policy = validate_choice(
            args.on_output_exists,
            OUTPUT_EXISTS_POLICIES,
            "--on-output-exists",
        )
        duplicate_name_policy = validate_choice(
            args.duplicate_name_policy,
            DUPLICATE_NAME_POLICIES,
            "--duplicate-name-policy",
        )

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

        warnings = build_warnings(records, total_pages, args.pages_per_file)
        errors: list[str] = []
        if args.output_dir and is_non_empty_directory(output_dir) and output_exists_policy == "fail":
            errors.append(f"Output directory already exists and is not empty: {output_dir}")

        config = JobConfig(
            sheet_path=sheet_path,
            pdf_path=pdf_path,
            pages_per_file=args.pages_per_file,
            naming_template=naming_template,
            output_dir=output_dir,
            name_column=detected_name_column,
            order_column=detected_order_column,
            output_exists_policy=output_exists_policy,
            duplicate_name_policy=duplicate_name_policy,
        )
        split_plan, duplicate_names, duplicate_rendered_filenames = plan_split_outputs(
            records,
            total_pages,
            config,
        )
        if duplicate_rendered_filenames and duplicate_name_policy == "fail":
            errors.append(
                "Duplicate rendered filenames were found: "
                + ", ".join(duplicate_rendered_filenames[:5])
            )

        status = "error" if errors else "warning" if warnings else "ok"
        preflight = FastCliPreflight(
            status=status,
            phase="validate",
            mode="split",
            merge_kind=None,
            can_proceed=not warnings and not errors,
            requires_confirmation=bool(warnings) and not errors,
            warnings=warnings,
            errors=errors,
            summary={
                "pdf_total_pages": total_pages,
                "sheet_record_count": len(records),
                "expected_output_files": min(
                    len(records),
                    (total_pages + args.pages_per_file - 1) // args.pages_per_file,
                ),
                "name_column": detected_name_column,
                "order_column": detected_order_column,
                "output_dir": str(output_dir),
                "duplicate_names": duplicate_names,
                "duplicate_rendered_filenames": duplicate_rendered_filenames,
                "duplicate_name_policy": duplicate_name_policy,
                "output_exists_policy": output_exists_policy,
                "planned_output_files": [str(entry.final_path) for entry in split_plan],
            },
        )
        context = {
            "kind": "split",
            "config": config,
            "records": records,
            "total_pages": total_pages,
            "warnings": warnings,
            "split_plan": split_plan,
        }
        return preflight, context
    except SystemExit as exc:
        return (
            FastCliPreflight(
                status="error",
                phase="validate",
                mode="split",
                merge_kind=None,
                can_proceed=False,
                requires_confirmation=False,
                warnings=[],
                errors=[str(exc)],
                summary={},
            ),
            None,
        )
    except Exception as exc:
        return build_preflight_exception("split", None, "Failed to inspect split inputs", exc)


def build_simple_merge_fast_cli_preflight(
    args: argparse.Namespace,
) -> tuple[FastCliPreflight, dict[str, object] | None]:
    try:
        if not args.first_pdf_path or not args.second_pdf_path:
            raise SystemExit("Simple merge mode requires --first-pdf-path and --second-pdf-path.")

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
        output_exists_policy = validate_choice(
            args.on_output_exists,
            OUTPUT_EXISTS_POLICIES,
            "--on-output-exists",
        )
        output_path = (
            build_default_merge_output_path(first_pdf_path)
            if not args.output_path
            else normalize_merge_output_path(
                Path(args.output_path).expanduser(),
                build_merge_output_filename(first_pdf_path),
            )
        )
        explicit_output_conflict = False
        if args.output_path:
            requested_output = Path(args.output_path).expanduser()
            explicit_output_conflict = detect_simple_merge_output_conflict(
                requested_output,
                build_merge_output_filename(first_pdf_path),
            )

        warnings: list[str] = []
        errors: list[str] = []
        if explicit_output_conflict and output_exists_policy == "fail":
            errors.append(f"Output file already exists: {output_path}")

        first_total_pages = get_pdf_page_count(first_pdf_path)
        second_total_pages = get_pdf_page_count(second_pdf_path)
        final_output_path, planned_action = resolve_output_target(output_path, output_exists_policy, set())
        preflight = FastCliPreflight(
            status="error" if errors else "warning" if warnings else "ok",
            phase="validate",
            mode="merge",
            merge_kind="simple",
            can_proceed=not warnings and not errors,
            requires_confirmation=bool(warnings) and not errors,
            warnings=warnings,
            errors=errors,
            summary={
                "first_pdf_pages": first_total_pages,
                "second_pdf_pages": second_total_pages,
                "output_path": str(final_output_path),
                "output_conflict": explicit_output_conflict,
                "output_exists_policy": output_exists_policy,
                "planned_output_file": str(final_output_path),
                "planned_action": planned_action,
            },
        )
        context = {
            "kind": "merge-simple",
            "config": MergeConfig(
                first_pdf_path=first_pdf_path,
                second_pdf_path=second_pdf_path,
                output_path=final_output_path,
                merge_order="first-second",
                output_exists_policy=output_exists_policy,
            ),
            "first_total_pages": first_total_pages,
            "second_total_pages": second_total_pages,
            "planned_action": planned_action,
        }
        return preflight, context
    except SystemExit as exc:
        return (
            FastCliPreflight(
                status="error",
                phase="validate",
                mode="merge",
                merge_kind="simple",
                can_proceed=False,
                requires_confirmation=False,
                warnings=[],
                errors=[str(exc)],
                summary={},
            ),
            None,
        )
    except Exception as exc:
        return build_preflight_exception("merge", "simple", "Failed to inspect merge inputs", exc)


def build_batch_merge_fast_cli_preflight(
    args: argparse.Namespace,
) -> tuple[FastCliPreflight, dict[str, object] | None]:
    try:
        if not args.batch_input_dir or not args.fixed_pdf_path:
            raise SystemExit("Batch merge mode requires --batch-input-dir and --fixed-pdf-path.")

        input_dir = validate_existing_directory_path(
            Path(args.batch_input_dir).expanduser(),
            "batch input folder",
        )
        fixed_pdf_path = validate_existing_file_path(
            Path(args.fixed_pdf_path).expanduser(),
            {".pdf"},
            "fixed PDF file",
        )
        output_exists_policy = validate_choice(
            args.on_output_exists,
            OUTPUT_EXISTS_POLICIES,
            "--on-output-exists",
        )
        input_pdfs = ensure_batch_input_pdfs_exist(input_dir)
        output_dir = (
            Path(args.batch_output_dir).expanduser()
            if args.batch_output_dir
            else build_default_batch_merge_output_dir(input_dir)
        )
        output_dir_conflict = bool(args.batch_output_dir and is_non_empty_directory(output_dir))
        warnings: list[str] = []
        errors: list[str] = []
        if output_dir_conflict and output_exists_policy == "fail":
            errors.append(f"Output directory already exists and is not empty: {output_dir}")

        fixed_pdf_pages = get_pdf_page_count(fixed_pdf_path)
        config = BatchMergeConfig(
            input_dir=input_dir,
            fixed_pdf_path=fixed_pdf_path,
            merge_order=args.merge_order,
            output_dir=output_dir,
            output_exists_policy=output_exists_policy,
        )
        batch_plan = plan_batch_outputs(config)
        preflight = FastCliPreflight(
            status="error" if errors else "warning" if warnings else "ok",
            phase="validate",
            mode="merge",
            merge_kind="batch",
            can_proceed=not warnings and not errors,
            requires_confirmation=bool(warnings) and not errors,
            warnings=warnings,
            errors=errors,
            summary={
                "input_dir": str(input_dir),
                "input_pdf_count": len(input_pdfs),
                "fixed_pdf_pages": fixed_pdf_pages,
                "merge_order": args.merge_order,
                "output_dir": str(output_dir),
                "output_dir_conflict": output_dir_conflict,
                "output_exists_policy": output_exists_policy,
                "planned_output_files": [str(entry.final_path) for entry in batch_plan],
            },
        )
        context = {
            "kind": "merge-batch",
            "config": config,
            "batch_plan": batch_plan,
        }
        return preflight, context
    except SystemExit as exc:
        return (
            FastCliPreflight(
                status="error",
                phase="validate",
                mode="merge",
                merge_kind="batch",
                can_proceed=False,
                requires_confirmation=False,
                warnings=[],
                errors=[str(exc)],
                summary={},
            ),
            None,
        )
    except Exception as exc:
        return build_preflight_exception("merge", "batch", "Failed to inspect batch merge inputs", exc)


def execute_fast_cli_context(
    context: dict[str, object] | None,
    preflight: FastCliPreflight,
) -> FastCliPreflight:
    if not context:
        return preflight

    if context["kind"] == "split":
        config = context["config"]
        records = context["records"]
        total_pages = context["total_pages"]
        warnings = context["warnings"]
        split_plan = context["split_plan"]
        assert isinstance(config, JobConfig)
        assert isinstance(records, list)
        assert isinstance(total_pages, int)
        assert isinstance(warnings, list)
        assert isinstance(split_plan, list)
        result = split_pdf_named(
            config,
            records,
            total_pages,
            show_progress_bar=False,
            split_plan=split_plan,
        )
        write_report(config, total_pages, len(records), warnings, result)
        return FastCliPreflight(
            status="ok",
            phase="execute",
            mode="split",
            merge_kind=None,
            can_proceed=True,
            requires_confirmation=False,
            warnings=preflight.warnings,
            errors=[],
            summary=preflight.summary,
            result={
                "written_files": result.written,
                "output_files": [str(path) for path in result.output_files],
                "overwritten_files": [str(path) for path in result.overwritten_files],
                "renamed_files": [str(path) for path in result.renamed_files],
                "skipped_existing_files": [str(path) for path in result.skipped_existing_files],
                "skipped_names": result.skipped_names,
                "unwritten_chunks": result.skipped_chunks,
                "report_path": str(config.output_dir / "split_report.txt"),
            },
        )

    if context["kind"] == "merge-simple":
        config = context["config"]
        first_total_pages = context["first_total_pages"]
        second_total_pages = context["second_total_pages"]
        planned_action = context["planned_action"]
        assert isinstance(config, MergeConfig)
        assert isinstance(first_total_pages, int)
        assert isinstance(second_total_pages, int)
        assert isinstance(planned_action, str)
        result = merge_pdf_files(config)
        write_merge_report(config, first_total_pages, second_total_pages, result)
        return FastCliPreflight(
            status="ok",
            phase="execute",
            mode="merge",
            merge_kind="simple",
            can_proceed=True,
            requires_confirmation=False,
            warnings=preflight.warnings,
            errors=[],
            summary=preflight.summary,
            result={
                "output_file": str(result.output_path),
                "total_pages": result.total_pages,
                "overwritten_files": [
                    str(path)
                    for path in (
                        result.overwritten_files
                        or ([result.output_path] if planned_action == "overwrite" else [])
                    )
                ],
                "renamed_files": [
                    str(path)
                    for path in (
                        result.renamed_files
                        or ([result.output_path] if planned_action == "rename" else [])
                    )
                ],
                "skipped_existing_files": [str(path) for path in result.skipped_existing_files],
                "report_path": str(result.output_path.parent / "merge_report.txt"),
            },
        )

    config = context["config"]
    batch_plan = context["batch_plan"]
    assert isinstance(config, BatchMergeConfig)
    assert isinstance(batch_plan, list)
    result = merge_pdf_folder(config, show_progress_bar=False, batch_plan=batch_plan)
    write_batch_merge_report(config, result)
    return FastCliPreflight(
        status="ok",
        phase="execute",
        mode="merge",
        merge_kind="batch",
        can_proceed=True,
        requires_confirmation=False,
        warnings=preflight.warnings,
        errors=[],
        summary=preflight.summary,
        result={
            "written_files": result.written,
            "output_files": [str(path) for path in result.output_files],
            "overwritten_files": [str(path) for path in result.overwritten_files],
            "renamed_files": [str(path) for path in result.renamed_files],
            "skipped_existing_files": [str(path) for path in result.skipped_existing_files],
            "report_path": str(config.output_dir / "merge_report.txt"),
        },
    )


def build_preflight_exception(
    mode: str,
    merge_kind: str | None,
    message: str,
    exc: Exception,
) -> tuple[FastCliPreflight, None]:
    return (
        FastCliPreflight(
            status="error",
            phase="validate",
            mode=mode,
            merge_kind=merge_kind,
            can_proceed=False,
            requires_confirmation=False,
            warnings=[],
            errors=[f"{message}: {exc}"],
            summary={},
        ),
        None,
    )


def emit_fast_cli_error(args: argparse.Namespace, message: str) -> int:
    preflight = FastCliPreflight(
        status="error",
        phase="validate",
        mode=args.mode or "unknown",
        merge_kind=args.merge_kind if args.mode == "merge" else None,
        can_proceed=False,
        requires_confirmation=False,
        warnings=[],
        errors=[message],
        summary={},
    )
    return emit_fast_cli_result(args, preflight, exit_code=1)


def emit_fast_cli_result(
    args: argparse.Namespace,
    preflight: FastCliPreflight,
    exit_code: int,
) -> int:
    if args.json:
        payload = fast_cli_preflight_to_dict(preflight)
        print(json.dumps(payload, ensure_ascii=True))
        return exit_code

    stream = sys.stderr if preflight.status != "ok" else sys.stdout
    print(render_fast_cli_preflight(preflight), file=stream)
    return exit_code


def fast_cli_preflight_to_dict(preflight: FastCliPreflight) -> dict[str, object]:
    return {
        "status": preflight.status,
        "phase": preflight.phase,
        "mode": preflight.mode,
        "merge_kind": preflight.merge_kind,
        "can_proceed": preflight.can_proceed,
        "requires_confirmation": preflight.requires_confirmation,
        "warnings": preflight.warnings,
        "errors": preflight.errors,
        "summary": preflight.summary,
        "result": preflight.result,
    }


def render_fast_cli_preflight(preflight: FastCliPreflight) -> str:
    lines = [
        "------------------------------------------------------------",
        "Fast CLI Summary",
        "------------------------------------------------------------",
        f"Phase                   : {preflight.phase}",
        f"Status                  : {preflight.status}",
        f"Mode                    : {preflight.mode}",
    ]
    if preflight.merge_kind:
        lines.append(f"Merge kind              : {preflight.merge_kind}")
    for key, value in preflight.summary.items():
        label = key.replace("_", " ").title()
        lines.append(f"{label:<24}: {value}")
    if preflight.warnings:
        lines.append("")
        lines.append("Warnings:")
        lines.extend(f"- {warning}" for warning in preflight.warnings)
    if preflight.errors:
        lines.append("")
        lines.append("Errors:")
        lines.extend(f"- {error}" for error in preflight.errors)
    if preflight.requires_confirmation:
        lines.append("")
        lines.append("Execution is blocked until you rerun with --confirm.")
    if preflight.phase == "execute" and preflight.result:
        lines.append("")
        lines.append("Result:")
        for key, value in preflight.result.items():
            label = key.replace("_", " ").title()
            lines.append(f"{label:<24}: {value}")
    return "\n".join(lines)


def run_startup_checks(
    simulated_missing: list[str] | None = None,
    interactive: bool = True,
    verbose: bool = True,
) -> None:
    missing = find_missing_dependencies(simulated_missing=simulated_missing)
    if missing:
        if verbose:
            print("Startup check found a missing required library:")
            for module_name in missing:
                print(f"- {module_name}")

        if not interactive:
            message = (
                "Required dependency is missing. Run 'Setup PDF Editor.command' or install manually with: "
                f"{sys.executable} -m pip install {' '.join(missing)}"
            )
            if verbose:
                print("\nPlease run 'Setup PDF Editor.command' or install manually with:")
                print(f"{sys.executable} -m pip install {' '.join(missing)}")
            raise SystemExit(message)

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

    if verbose:
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


def prompt_merge_kind() -> str:
    while True:
        raw = input(
            "Which merge function do you want to use?\n"
            "1. Simple Merge\n"
            "2. Batch Merge\n"
            "> "
        ).strip().lower()
        if raw in {"1", "simple", "simple merge"}:
            return "simple"
        if raw in {"2", "batch", "batch merge"}:
            return "batch"
        print("Please choose 1 for Simple Merge or 2 for Batch Merge.")


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


def prompt_batch_merge_order() -> str:
    while True:
        raw = input(
            "\n[3/4] Which merge order do you want?\n"
            "1. Split PDF first, fixed PDF second\n"
            "2. Fixed PDF first, split PDF second\n"
            "> "
        ).strip().lower()
        if raw in {"1", "split-first", "split first"}:
            return "split-first"
        if raw in {"2", "fixed-first", "fixed first"}:
            return "fixed-first"
        print("Please choose 1 for split-first or 2 for fixed-first.")


def prompt_batch_merge_output_dir(message: str, input_dir: Path) -> Path:
    raw = input(message)
    path = parse_path_input(raw)
    if path:
        return path
    output_dir = build_default_batch_merge_output_dir(input_dir)
    print("\nNo output folder was provided.")
    print('Default folder name will be: "Batch Merged PDF"')
    print(f"I will create this folder automatically:\n{output_dir}")
    return output_dir


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


def validate_existing_directory_path(path: Path, label: str = "folder") -> Path:
    if not path.exists():
        raise SystemExit(f"The {label} was not found: {path}")
    if not path.is_dir():
        raise SystemExit(f"The {label} is not a folder: {path}")
    return path


def prompt_existing_directory(message: str) -> Path:
    while True:
        raw = input(f"{message}\n> ")
        path = parse_path_input(raw)
        if not path:
            print("Please enter a valid path.")
            continue
        if not path.exists():
            print("That folder was not found. Please try again.")
            continue
        if not path.is_dir():
            print("That path is not a folder. Please try again.")
            continue
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


def validate_choice(value: str, allowed: set[str], label: str) -> str:
    if value not in allowed:
        allowed_values = ", ".join(sorted(allowed))
        raise SystemExit(f"{label} must be one of: {allowed_values}")
    return value


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


def append_marker_to_filename(filename: str, marker: str) -> str:
    path = Path(filename)
    return f"{sanitize_filename(f'{path.stem} [{marker}]')}{path.suffix}"


def find_duplicate_names(records: list[InputRecord]) -> list[str]:
    return sorted(name for name, count in Counter(record.name for record in records).items() if count > 1)


def find_duplicate_rendered_filenames(filenames: list[str]) -> list[str]:
    return sorted(name for name, count in Counter(filenames).items() if count > 1)


def is_non_empty_directory(path: Path) -> bool:
    return path.exists() and path.is_dir() and any(path.iterdir())


def detect_simple_merge_output_conflict(requested_output: Path, default_filename: str) -> bool:
    if requested_output.exists() and requested_output.is_dir():
        return (requested_output / default_filename).exists()
    if requested_output.suffix.casefold() != ".pdf":
        requested_output = requested_output.with_suffix(".pdf")
    return requested_output.exists()


def build_warnings(records: list[InputRecord], total_pages: int, pages_per_file: int) -> list[str]:
    warnings: list[str] = []
    chunk_count = (total_pages + pages_per_file - 1) // pages_per_file
    if len(records) != chunk_count:
        warnings.append(
            f"Sheet record count ({len(records)}) does not match PDF chunk count ({chunk_count})."
        )
    return warnings


def build_unique_candidate(path: Path, reserved_paths: set[Path]) -> Path:
    if path not in reserved_paths and not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        if candidate not in reserved_paths and not candidate.exists():
            return candidate
        counter += 1


def resolve_output_target(
    requested_path: Path,
    output_exists_policy: str,
    reserved_paths: set[Path],
) -> tuple[Path, str]:
    if requested_path in reserved_paths:
        return build_unique_candidate(requested_path, reserved_paths), "rename"
    if not requested_path.exists():
        return requested_path, "write"
    if output_exists_policy == "overwrite":
        return requested_path, "overwrite"
    if output_exists_policy in {"rename", "continue"}:
        return build_unique_candidate(requested_path, reserved_paths), "rename"
    return requested_path, "skip"


def plan_split_outputs(
    records: list[InputRecord],
    total_pages: int,
    config: JobConfig,
) -> tuple[list[PlannedSplitOutput], list[str], list[str]]:
    chunk_count = (total_pages + config.pages_per_file - 1) // config.pages_per_file
    limit = min(len(records), chunk_count)
    planned_records = records[:limit]
    base_filenames = [build_output_filename(record.name, config.naming_template) for record in planned_records]
    duplicate_rendered_filenames = find_duplicate_rendered_filenames(base_filenames)
    duplicate_name_lookup = set(duplicate_rendered_filenames)
    duplicate_names = sorted(
        {
            record.name
            for record, filename in zip(planned_records, base_filenames, strict=False)
            if filename in duplicate_name_lookup
        }
    )

    requested_filenames: list[str] = []
    for record, filename in zip(planned_records, base_filenames, strict=False):
        requested_filename = filename
        if filename in duplicate_name_lookup:
            if config.duplicate_name_policy == "append-row-number":
                requested_filename = append_marker_to_filename(filename, f"row-{record.index}")
            elif config.duplicate_name_policy == "append-order":
                marker = (
                    f"order-{record.order}"
                    if config.order_column
                    else f"row-{record.index}"
                )
                requested_filename = append_marker_to_filename(filename, marker)
        requested_filenames.append(requested_filename)

    reserved_paths: set[Path] = set()
    plan: list[PlannedSplitOutput] = []
    for record, requested_filename in zip(planned_records, requested_filenames, strict=False):
        requested_path = config.output_dir / requested_filename
        final_path, action = resolve_output_target(
            requested_path,
            config.output_exists_policy,
            reserved_paths,
        )
        reserved_paths.add(final_path)
        plan.append(
            PlannedSplitOutput(
                record=record,
                requested_filename=requested_filename,
                final_path=final_path,
                action=action,
            )
        )
    return plan, duplicate_names, duplicate_rendered_filenames


def plan_batch_outputs(config: BatchMergeConfig) -> list[PlannedBatchOutput]:
    reserved_paths: set[Path] = set()
    plan: list[PlannedBatchOutput] = []
    for input_pdf_path in collect_batch_input_pdfs(config.input_dir):
        requested_path = config.output_dir / input_pdf_path.name
        final_path, action = resolve_output_target(
            requested_path,
            config.output_exists_policy,
            reserved_paths,
        )
        reserved_paths.add(final_path)
        plan.append(
            PlannedBatchOutput(
                input_pdf_path=input_pdf_path,
                final_path=final_path,
                action=action,
            )
        )
    return plan


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


def split_pdf_named(
    config: JobConfig,
    records: list[InputRecord],
    total_pages: int,
    show_progress_bar: bool = True,
    split_plan: list[PlannedSplitOutput] | None = None,
) -> SplitResult:
    pdf_reader, pdf_writer = load_pdf_tools()
    reader = pdf_reader(str(config.pdf_path))
    config.output_dir.mkdir(parents=True, exist_ok=True)

    chunk_count = (total_pages + config.pages_per_file - 1) // config.pages_per_file
    limit = min(len(records), chunk_count)
    output_files: list[Path] = []
    overwritten_files: list[Path] = []
    renamed_files: list[Path] = []
    skipped_existing_files: list[Path] = []
    plan = split_plan or plan_split_outputs(records, total_pages, config)[0]

    for idx in range(limit):
        start = idx * config.pages_per_file
        end = min(start + config.pages_per_file, total_pages)

        writer = pdf_writer()
        for page_number in range(start, end):
            writer.add_page(reader.pages[page_number])

        planned = plan[idx]
        output_path = planned.final_path
        with output_path.open("wb") as handle:
            writer.write(handle)
        output_files.append(output_path)
        if planned.action == "overwrite":
            overwritten_files.append(output_path)
        elif planned.action == "rename":
            renamed_files.append(output_path)
        elif planned.action == "skip":
            skipped_existing_files.append(output_path)
        if show_progress_bar:
            print_progress(idx + 1, limit)

    if show_progress_bar:
        print("")
    return SplitResult(
        written=limit,
        skipped_names=max(0, len(records) - limit),
        skipped_chunks=max(0, chunk_count - limit),
        output_files=output_files,
        overwritten_files=overwritten_files,
        renamed_files=renamed_files,
        skipped_existing_files=skipped_existing_files,
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
    total_pages = get_pdf_page_count(config.first_pdf_path) + get_pdf_page_count(config.second_pdf_path)
    output_path, action = resolve_output_target(config.output_path, config.output_exists_policy, set())
    output_path = merge_two_pdf_paths(
        config.first_pdf_path,
        config.second_pdf_path,
        output_path,
        config.merge_order,
    )
    return MergeResult(
        output_path=output_path,
        total_pages=total_pages,
        overwritten_files=[output_path] if action == "overwrite" else [],
        renamed_files=[output_path] if action == "rename" else [],
        skipped_existing_files=[],
    )


def show_batch_merge_summary(config: BatchMergeConfig) -> None:
    input_pdfs = ensure_batch_input_pdfs_exist(config.input_dir)
    fixed_total_pages = get_pdf_page_count(config.fixed_pdf_path)
    merged_pages = sum(get_pdf_page_count(path) + fixed_total_pages for path in input_pdfs)
    print("\n------------------------------------------------------------")
    print("Review this summary before I start:")
    print("------------------------------------------------------------")
    print(f"Batch input folder      : {config.input_dir}")
    print(f"Found split PDF files   : {len(input_pdfs)}")
    print(f"Fixed PDF               : {config.fixed_pdf_path}")
    print(f"Fixed PDF total pages   : {fixed_total_pages}")
    print(f"Merge order             : {config.merge_order}")
    print(f"Batch output folder     : {config.output_dir}")
    print(f"Expected output files   : {len(input_pdfs)}")
    print(f"Expected merged pages   : {merged_pages}")


def merge_pdf_folder(
    config: BatchMergeConfig,
    show_progress_bar: bool = True,
    batch_plan: list[PlannedBatchOutput] | None = None,
) -> BatchMergeResult:
    input_pdfs = ensure_batch_input_pdfs_exist(config.input_dir)
    config.output_dir.mkdir(parents=True, exist_ok=True)
    fixed_total_pages = get_pdf_page_count(config.fixed_pdf_path)
    output_files: list[Path] = []
    overwritten_files: list[Path] = []
    renamed_files: list[Path] = []
    skipped_existing_files: list[Path] = []
    plan = batch_plan or plan_batch_outputs(config)

    for index, planned in enumerate(plan, start=1):
        output_path = merge_two_pdf_paths(
            planned.input_pdf_path,
            config.fixed_pdf_path,
            planned.final_path,
            config.merge_order,
        )
        output_files.append(output_path)
        if planned.action == "overwrite":
            overwritten_files.append(output_path)
        elif planned.action == "rename":
            renamed_files.append(output_path)
        elif planned.action == "skip":
            skipped_existing_files.append(output_path)
        if show_progress_bar:
            print_progress(index, len(plan))

    if show_progress_bar:
        print("")
    total_pages_per_file = fixed_total_pages
    if input_pdfs:
        total_pages_per_file += get_pdf_page_count(input_pdfs[0])
    return BatchMergeResult(
        written=len(output_files),
        total_pages_per_file=total_pages_per_file,
        output_files=output_files,
        overwritten_files=overwritten_files,
        renamed_files=renamed_files,
        skipped_existing_files=skipped_existing_files,
    )


def merge_two_pdf_paths(
    first_pdf_path: Path,
    second_pdf_path: Path,
    output_path: Path,
    merge_order: str,
) -> Path:
    pdf_reader, pdf_writer = load_pdf_tools()
    first_reader = pdf_reader(str(first_pdf_path))
    second_reader = pdf_reader(str(second_pdf_path))
    writer = pdf_writer()

    readers = (first_reader, second_reader)
    if merge_order == "fixed-first":
        readers = (second_reader, first_reader)

    for reader in readers:
        for page in reader.pages:
            writer.add_page(page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as handle:
        writer.write(handle)
    return output_path


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
    return build_default_merge_output_dir(first_pdf_path) / build_merge_output_filename(first_pdf_path)


def build_default_batch_merge_output_dir(input_dir: Path) -> Path:
    return ensure_unique_directory_path(input_dir.with_name("Batch Merged PDF"))


def collect_batch_input_pdfs(input_dir: Path) -> list[Path]:
    return sorted(
        [
            path
            for path in input_dir.iterdir()
            if path.is_file() and path.suffix.casefold() == ".pdf"
        ],
        key=lambda path: path.name.casefold(),
    )


def ensure_batch_input_pdfs_exist(input_dir: Path) -> list[Path]:
    input_pdfs = collect_batch_input_pdfs(input_dir)
    if not input_pdfs:
        raise SystemExit("The batch input folder does not contain any PDF files.")
    return input_pdfs


def normalize_merge_output_path(path: Path, default_filename: str) -> Path:
    if path.exists() and path.is_dir():
        return path / default_filename
    if path.suffix.casefold() != ".pdf":
        path = path.with_suffix(".pdf")
    return path


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
        f"Merge order: {config.merge_order}",
        f"Output file: {result.output_path}",
        "",
        f"First PDF total pages: {first_total_pages}",
        f"Second PDF total pages: {second_total_pages}",
        f"Total merged pages: {result.total_pages}",
    ]
    report_path.write_text("\n".join(lines), encoding="utf-8")


def write_batch_merge_report(config: BatchMergeConfig, result: BatchMergeResult) -> None:
    report_path = config.output_dir / "merge_report.txt"
    lines = [
        "PDF EDITOR BATCH MERGE REPORT",
        "",
        f"Batch input folder: {config.input_dir}",
        f"Fixed PDF path: {config.fixed_pdf_path}",
        f"Merge order: {config.merge_order}",
        f"Output folder: {config.output_dir}",
        f"Written files: {result.written}",
        "",
        "Output files:",
    ]
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


def show_merge_completion(result: MergeResult) -> None:
    print("------------------------------------------------------------")
    print("Done")
    print("------------------------------------------------------------")
    print(f"Merged PDF file         : {result.output_path}")
    print(f"Total merged pages      : {result.total_pages}")
    print(f"Output folder           : {result.output_path.parent}")
    print(f"Report file             : {result.output_path.parent / 'merge_report.txt'}")


def show_batch_merge_completion(config: BatchMergeConfig, result: BatchMergeResult) -> None:
    print("------------------------------------------------------------")
    print("Done")
    print("------------------------------------------------------------")
    print(f"Generated PDF files     : {result.written}")
    print(f"Output folder           : {config.output_dir}")
    print(f"Report file             : {config.output_dir / 'merge_report.txt'}")

    if result.output_files:
        print("\nExample files:")
        for path in result.output_files[:3]:
            print(f"- {path.name}")
