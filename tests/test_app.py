import argparse
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from pdf_editor.app import (
    MergeConfig,
    build_default_output_dir,
    build_default_merge_output_path,
    build_default_merge_output_dir,
    build_default_output_dir_label,
    build_merge_output_filename,
    build_output_filename,
    contains_name_placeholder,
    ensure_unique_directory_path,
    ensure_unique_path,
    find_missing_dependencies,
    merge_pdf_files,
    inspect_sheet,
    install_missing_dependencies,
    normalize_key,
    normalize_merge_output_path,
    parse_simulated_missing_deps,
    parse_path_input,
    pick_column,
    read_sheet_records,
    resolve_requested_column_name,
    run_non_interactive,
    sanitize_filename,
    sanitize_naming_template,
    strip_simulated_missing_args,
    validate_existing_file_path,
)


def create_minimal_xlsx(path: Path) -> None:
    with ZipFile(path, "w") as workbook:
        workbook.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
""",
        )
        workbook.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
""",
        )
        workbook.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
""",
        )
        workbook.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
""",
        )
        workbook.writestr(
            "xl/sharedStrings.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="6">
  <si><t>NO</t></si>
  <si><t>NAME</t></si>
  <si><t>1</t></si>
  <si><t>Alice Tan</t></si>
  <si><t>2</t></si>
  <si><t>Bob Lee</t></si>
</sst>
""",
        )
        workbook.writestr(
            "xl/worksheets/sheet1.xml",
            """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>2</v></c>
      <c r="B2" t="s"><v>3</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>4</v></c>
      <c r="B3" t="s"><v>5</v></c>
    </row>
  </sheetData>
</worksheet>
""",
        )


class AppTests(unittest.TestCase):
    def test_sanitize_filename_removes_illegal_characters(self) -> None:
        self.assertEqual(sanitize_filename('A/B:C*"D"?'), "A B C D")

    def test_build_output_filename_with_template(self) -> None:
        self.assertEqual(
            build_output_filename(
                "Alice Tan",
                "GD Pink Form - Letter of Offer ({Name}) 26-3-2026",
            ),
            "GD Pink Form - Letter of Offer (Alice Tan) 26-3-2026.pdf",
        )

    def test_build_output_filename_with_name_only_template(self) -> None:
        self.assertEqual(build_output_filename("Alice Tan", "{Name}"), "Alice Tan.pdf")

    def test_build_output_filename_supports_lowercase_placeholder(self) -> None:
        self.assertEqual(
            build_output_filename("Alice Tan", "Letter ({name}).pdf"),
            "Letter (Alice Tan).pdf",
        )

    def test_pick_column_ignores_case_and_spacing(self) -> None:
        fieldnames = ["Full Name", "Employee No"]
        self.assertEqual(pick_column(fieldnames, ["full_name"]), "Full Name")

    def test_parse_path_input_handles_quotes(self) -> None:
        path = parse_path_input('"/tmp/example file.csv"')
        self.assertEqual(path, Path("/tmp/example file.csv"))

    def test_normalize_key(self) -> None:
        self.assertEqual(normalize_key("Employee_Name"), "employeename")

    def test_find_missing_dependencies_reports_missing_module(self) -> None:
        def fake_loader(name: str) -> object:
            raise ModuleNotFoundError(name)

        self.assertEqual(find_missing_dependencies(fake_loader), ["pypdf"])

    def test_find_missing_dependencies_returns_empty_when_present(self) -> None:
        def fake_loader(name: str) -> object:
            return object()

        self.assertEqual(find_missing_dependencies(fake_loader), [])

    def test_find_missing_dependencies_respects_simulated_missing(self) -> None:
        def fake_loader(name: str) -> object:
            return object()

        self.assertEqual(
            find_missing_dependencies(fake_loader, simulated_missing=["pypdf"]),
            ["pypdf"],
        )

    def test_install_missing_dependencies_returns_true_on_success(self) -> None:
        calls: list[list[str]] = []

        def fake_installer(modules: list[str]) -> int:
            calls.append(modules)
            return 0

        self.assertTrue(install_missing_dependencies(["pypdf"], fake_installer))
        self.assertEqual(calls, [["pypdf"]])

    def test_install_missing_dependencies_returns_false_on_failure(self) -> None:
        def fake_installer(modules: list[str]) -> int:
            return 1

        self.assertFalse(install_missing_dependencies(["pypdf"], fake_installer))

    def test_ensure_unique_path_adds_counter(self) -> None:
        with TemporaryDirectory() as tmpdir:
            base = Path(tmpdir) / "Alice.pdf"
            base.write_text("x", encoding="utf-8")
            unique = ensure_unique_path(base)
            self.assertEqual(unique.name, "Alice (2).pdf")

    def test_ensure_unique_directory_path_adds_counter(self) -> None:
        with TemporaryDirectory() as tmpdir:
            base = Path(tmpdir) / "report_split_output"
            base.mkdir()
            unique = ensure_unique_directory_path(base)
            self.assertEqual(unique.name, "report_split_output (2)")

    def test_build_default_output_dir_uses_pdf_stem_for_name_only_template(self) -> None:
        pdf_path = Path("/tmp/EA Form (Updated).pdf")
        self.assertEqual(build_default_output_dir(pdf_path, "{Name}").name, "EA Form (Updated)")

    def test_build_default_output_dir_prefers_template_text(self) -> None:
        pdf_path = Path("/tmp/EA Form_removed.pdf")
        self.assertEqual(
            build_default_output_dir(
                pdf_path,
                "GD Pink Form - Letter of Offer ({Name}) 26-3-2026",
            ).name,
            "GD Pink Form - Letter of Offer 26-3-2026",
        )

    def test_sanitize_naming_template_strips_pdf_extension(self) -> None:
        self.assertEqual(
            sanitize_naming_template("GD Pink Form ({Name}).pdf"),
            "GD Pink Form ({Name})",
        )

    def test_contains_name_placeholder_accepts_case_and_spacing(self) -> None:
        self.assertTrue(contains_name_placeholder("Letter ({ Name })"))
        self.assertFalse(contains_name_placeholder("Letter (Employee)"))

    def test_build_default_output_dir_label_removes_empty_brackets(self) -> None:
        self.assertEqual(
            build_default_output_dir_label("Letter of Offer ({Name}) 26-3-2026"),
            "Letter of Offer 26-3-2026",
        )

    def test_build_default_merge_output_dir_uses_merged_pdf_name(self) -> None:
        pdf_path = Path("/tmp/source.pdf")
        self.assertEqual(build_default_merge_output_dir(pdf_path).name, "Merged PDF")

    def test_build_default_merge_output_path_uses_first_pdf_name(self) -> None:
        pdf_path = Path("/tmp/Offer Letter.pdf")
        self.assertEqual(
            build_default_merge_output_path(pdf_path),
            Path("/tmp/Merged PDF/Offer Letter.pdf"),
        )

    def test_build_merge_output_filename_follows_first_pdf_name(self) -> None:
        pdf_path = Path("/tmp/Offer Letter.pdf")
        self.assertEqual(build_merge_output_filename(pdf_path), "Offer Letter.pdf")

    def test_normalize_merge_output_path_uses_default_filename_for_directory(self) -> None:
        with TemporaryDirectory() as tmpdir:
            output_dir = Path(tmpdir)
            self.assertEqual(
                normalize_merge_output_path(output_dir, "Offer Letter.pdf"),
                output_dir / "Offer Letter.pdf",
            )

    def test_normalize_merge_output_path_adds_pdf_extension(self) -> None:
        output_path = Path("/tmp/merged_output")
        self.assertEqual(
            normalize_merge_output_path(output_path, "Offer Letter.pdf"),
            Path("/tmp/merged_output.pdf"),
        )

    def test_merge_pdf_files_writes_merged_output(self) -> None:
        class FakeReader:
            def __init__(self, path: str) -> None:
                self.pages = [f"{Path(path).stem}-page-1", f"{Path(path).stem}-page-2"]

        class FakeWriter:
            def __init__(self) -> None:
                self.pages: list[str] = []

            def add_page(self, page: str) -> None:
                self.pages.append(page)

            def write(self, handle) -> None:
                handle.write("\n".join(self.pages).encode("utf-8"))

        with TemporaryDirectory() as tmpdir:
            first_pdf = Path(tmpdir) / "first.pdf"
            second_pdf = Path(tmpdir) / "second.pdf"
            first_pdf.write_bytes(b"first")
            second_pdf.write_bytes(b"second")
            output_path = Path(tmpdir) / "merged.pdf"
            config = MergeConfig(
                first_pdf_path=first_pdf,
                second_pdf_path=second_pdf,
                output_path=output_path,
            )

            from unittest.mock import patch

            with patch("pdf_editor.app.load_pdf_tools", return_value=(FakeReader, FakeWriter)):
                result = merge_pdf_files(config)

            self.assertEqual(result.output_path, output_path)
            self.assertEqual(result.total_pages, 4)
            self.assertEqual(
                output_path.read_text(encoding="utf-8"),
                "first-page-1\nfirst-page-2\nsecond-page-1\nsecond-page-2",
            )

    def test_validate_existing_file_path_rejects_missing_file(self) -> None:
        with self.assertRaises(SystemExit) as context:
            validate_existing_file_path(Path("/tmp/does-not-exist.pdf"), {".pdf"}, "PDF file")
        self.assertIn("was not found", str(context.exception))

    def test_resolve_requested_column_name_matches_case_insensitively(self) -> None:
        self.assertEqual(
            resolve_requested_column_name(["Full Name", "Employee No"], "full_name", "name"),
            "Full Name",
        )

    def test_run_non_interactive_requires_split_inputs(self) -> None:
        args = argparse.Namespace(
            mode="split",
            sheet_path=None,
            pdf_path=None,
            pages_per_file=1,
            naming_template="{Name}",
            output_dir=None,
            name_column=None,
            order_column=None,
            first_pdf_path=None,
            second_pdf_path=None,
            output_path=None,
        )
        from unittest.mock import patch

        with patch("pdf_editor.app.run_startup_checks"):
            with self.assertRaises(SystemExit) as context:
                run_non_interactive(args)
        self.assertIn("--sheet-path and --pdf-path", str(context.exception))

    def test_run_non_interactive_requires_merge_inputs(self) -> None:
        args = argparse.Namespace(
            mode="merge",
            sheet_path=None,
            pdf_path=None,
            pages_per_file=1,
            naming_template="{Name}",
            output_dir=None,
            name_column=None,
            order_column=None,
            first_pdf_path=None,
            second_pdf_path=None,
            output_path=None,
        )
        from unittest.mock import patch

        with patch("pdf_editor.app.run_startup_checks"):
            with self.assertRaises(SystemExit) as context:
                run_non_interactive(args)
        self.assertIn("--first-pdf-path and --second-pdf-path", str(context.exception))

    def test_parse_simulated_missing_deps(self) -> None:
        self.assertEqual(parse_simulated_missing_deps("pypdf, other "), ["pypdf", "other"])

    def test_strip_simulated_missing_args(self) -> None:
        self.assertEqual(
            strip_simulated_missing_args(["--simulate-missing-deps", "pypdf", "--version"]),
            ["--version"],
        )
        self.assertEqual(
            strip_simulated_missing_args(["--simulate-missing-deps=pypdf", "--version"]),
            ["--version"],
        )

    def test_inspect_sheet_supports_xlsx(self) -> None:
        with TemporaryDirectory() as tmpdir:
            xlsx_path = Path(tmpdir) / "employees.xlsx"
            create_minimal_xlsx(xlsx_path)
            self.assertEqual(inspect_sheet(xlsx_path), ["NO", "NAME"])

    def test_read_sheet_records_supports_xlsx(self) -> None:
        with TemporaryDirectory() as tmpdir:
            xlsx_path = Path(tmpdir) / "employees.xlsx"
            create_minimal_xlsx(xlsx_path)
            _, records, name_column, order_column = read_sheet_records(xlsx_path)
            self.assertEqual(name_column, "NAME")
            self.assertEqual(order_column, "NO")
            self.assertEqual([record.name for record in records], ["Alice Tan", "Bob Lee"])


if __name__ == "__main__":
    unittest.main()
