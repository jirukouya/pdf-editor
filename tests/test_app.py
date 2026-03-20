import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from pdf_editor.app import (
    build_default_output_dir,
    build_output_filename,
    ensure_unique_directory_path,
    ensure_unique_path,
    find_missing_dependencies,
    inspect_sheet,
    install_missing_dependencies,
    normalize_key,
    parse_simulated_missing_deps,
    parse_path_input,
    pick_column,
    read_sheet_records,
    sanitize_filename,
    sanitize_suffix,
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

    def test_build_output_filename_with_suffix(self) -> None:
        self.assertEqual(
            build_output_filename("Alice Tan", "EA Form Revised"),
            "Alice Tan - EA Form Revised.pdf",
        )

    def test_build_output_filename_without_suffix(self) -> None:
        self.assertEqual(build_output_filename("Alice Tan", ""), "Alice Tan.pdf")

    def test_build_output_filename_strips_pdf_from_suffix(self) -> None:
        self.assertEqual(
            build_output_filename("Alice Tan", "Test.pdf"),
            "Alice Tan - Test.pdf",
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

    def test_build_default_output_dir_uses_pdf_stem(self) -> None:
        pdf_path = Path("/tmp/EA Form (Updated).pdf")
        self.assertEqual(build_default_output_dir(pdf_path, "").name, "EA Form (Updated)")

    def test_build_default_output_dir_prefers_user_suffix(self) -> None:
        pdf_path = Path("/tmp/EA Form_removed.pdf")
        self.assertEqual(build_default_output_dir(pdf_path, "Test").name, "Test")

    def test_sanitize_suffix_strips_pdf_extension(self) -> None:
        self.assertEqual(sanitize_suffix("EA Form (Updated).pdf"), "EA Form (Updated)")

    def test_parse_simulated_missing_deps(self) -> None:
        self.assertEqual(parse_simulated_missing_deps("pypdf, other "), ["pypdf", "other"])

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
