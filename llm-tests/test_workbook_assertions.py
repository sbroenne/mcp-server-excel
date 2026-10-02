"""Deterministic verification tests; run with unittest, without an external model."""

from __future__ import annotations

import copy
import hashlib
import json
import os
import subprocess
import tempfile
import unittest
import uuid
from pathlib import Path

from workbook_assertions import (
    assert_chart, assert_combined_slicers, assert_pivot_slicer, assert_table_slicers,
    read_saved_workbook,
)


class SavedWorkbookChecks(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix="excel-outcome-check-")
        self.path = Path(self.directory.name) / "result.xlsx"
        self.env = {**os.environ, "EXCELMCP_CLI_PIPE": f"outcome-check-{uuid.uuid4().hex}"}
        self.exe = Path(__file__).resolve().parents[1] / "src" / "ExcelMcp.CLI" / "bin" / "Release" / "net10.0-windows" / "excelcli.exe"
        self.session = None
        self.addCleanup(self._cleanup)
        self.session = self._cli("session", "create", str(self.path))["sessionId"]

    def _cli(self, *args):
        result = subprocess.run(
            [str(self.exe), "-q", *args], env=self.env,
            capture_output=True, text=True, encoding="utf-8", timeout=180, check=True,
        )
        data = json.loads(result.stdout)
        self.assertIsNot(data.get("success"), False, data)
        return data

    def _command(self, group, action, *args):
        return self._cli(group, action, "--session", self.session, *args)

    def _cleanup(self):
        if self.session:
            self._command("session", "close")
            self.session = None
        self._cli("service", "stop")
        self.directory.cleanup()

    def _write(self, cells, address):
        self._command("range", "set-values", "--sheet", "Sheet1", "--range", address,
                      "--values", json.dumps(cells), "--overwrite-policy", "allow")

    def _inspect(self, source_range="A1:E9"):
        self._command("session", "close", "--save")
        self.session = None
        before = hashlib.sha256(self.path.read_bytes()).digest()
        result = read_saved_workbook(str(self.path), source_range)
        self.assertEqual(before, hashlib.sha256(self.path.read_bytes()).digest(), "Inspection changed the saved file")
        return result

    def _reopen(self):
        self.session = self._cli("session", "open", str(self.path))["sessionId"]

    def test_used_range_reader_covers_cells_outside_the_default_window(self):
        self._write([["Far column"]], "Z3")
        self._command("session", "close", "--save")
        self.session = None
        before = hashlib.sha256(self.path.read_bytes()).digest()
        snapshot = read_saved_workbook(str(self.path), use_used_range=True, recalculate=True)
        sheet = snapshot["sheets"][0]
        self.assertEqual(sheet["sourceRow"], 3)
        self.assertEqual(sheet["sourceColumn"], 26)
        self.assertEqual(sheet["sourceValues"], [["Far column"]])
        self.assertEqual(before, hashlib.sha256(self.path.read_bytes()).digest())

    def test_in_memory_probe_recalculates_without_saving(self):
        self._write([[6]], "B3")
        self._command("range", "set-formulas", "--sheet", "Sheet1", "--range", "A3", "--formulas", '[["=B3*2"]]')
        self._command("session", "close", "--save")
        self.session = None
        before = hashlib.sha256(self.path.read_bytes()).digest()
        snapshot = read_saved_workbook(
            str(self.path), use_used_range=True, recalculate=True,
            probe_changes=[{"sheet": "Sheet1", "cell": "B3", "value": 7}],
        )
        self.assertEqual(snapshot["sheets"][0]["sourceValues"], [[14, 7]])
        self.assertEqual(before, hashlib.sha256(self.path.read_bytes()).digest())
        unchanged = read_saved_workbook(str(self.path), use_used_range=True, recalculate=True)
        self.assertEqual(unchanged["sheets"][0]["sourceValues"], [[12, 6]])

    def test_used_range_limit_rejects_10001_cells_without_changing_the_file(self):
        self._write([["Start"]], "A1")
        self._write([["Outside inspection limit"]], "A10001")
        self._command("session", "close", "--save")
        self.session = None
        before = hashlib.sha256(self.path.read_bytes()).digest()
        with self.assertRaisesRegex(AssertionError, "10000-cell"):
            read_saved_workbook(str(self.path), use_used_range=True)
        self.assertEqual(before, hashlib.sha256(self.path.read_bytes()).digest())

    def test_documented_batch_saves_only_a_complete_job(self):
        guide = self.exe.parents[5] / "docs" / "reference" / "workflows.md"
        content = guide.read_text(encoding="utf-8")
        recipe = content.split("```cli\n", 1)[1].split("\n```", 1)[0]
        environment = {**self.env, "PATH": f"{self.exe.parent}{os.pathsep}{self.env['PATH']}"}
        script = Path(self.directory.name) / "batch-example.ps1"
        script.write_text("param([string]$path)\n" + recipe, encoding="utf-8")
        refused = subprocess.run(
            ["pwsh", "-NoProfile", "-File", str(script), "-path", str(self.path)],
            env=environment, cwd=self.directory.name, capture_output=True,
            text=True, encoding="utf-8", timeout=180, check=False,
        )
        self.assertNotEqual(refused.returncode, 0)
        self.assertIn("Reuse the existing session", refused.stderr)
        self.assertEqual(self._cli("session", "list")["sessions"][0]["sessionId"], self.session)
        self._command("sheet", "create", "--sheet", "Sales")
        for fail in (False, True):
            with self.subTest(fail=fail):
                self._write([["Original1"], ["Original2"]], "A1:A2")
                self._command("range", "clear-contents", "--sheet", "Sales", "--range", "A1:B2")
                self._command("session", "close", "--save")
                self.session = None
                code = recipe
                if fail:
                    code = code.replace('"sheetName":"Sales","rangeAddress":"B2"',
                                        '"sheetName":"MissingSheet","rangeAddress":"B2"')
                script.write_text("param([string]$path)\n" + code, encoding="utf-8")
                result = subprocess.run(
                    ["pwsh", "-NoProfile", "-File", str(script), "-path", str(self.path)],
                    env=environment, cwd=self.directory.name, capture_output=True,
                    text=True, encoding="utf-8", timeout=180, check=False,
                )
                self.assertEqual(result.returncode == 0, not fail, result.stderr)
                self.assertEqual(self._cli("session", "list")["sessions"], [])
                self._reopen()
                values = self._command("range", "get-values", "--sheet", "Sheet1", "--range", "A1:A2")["values"]
                self.assertEqual(values, [["Original1"], ["Original2"]])
                sales = self._command("range", "get-values", "--sheet", "Sales", "--range", "A1:B2")["values"]
                self.assertEqual(sales, [[None, None], [None, None]] if fail
                                 else [["Product", "Amount"], ["Widget", 1250]])

    def test_snapshot_distinguishes_formulas_formats_and_query_identity(self):
        self._write([["X"], [7]], "A1:A2")
        self._command("range", "set-formulas", "--sheet", "Sheet1", "--range", "B2", "--formulas", '[["=A2*3"]]')
        self._command("range", "set-number-format", "--sheet", "Sheet1", "--range", "B2", "--format-code", "0.00")
        self._command("powerquery", "create", "--query-name", "A",
                      "--load-destination", "connection-only", "--m-code", "#table(type table [X=number], {{1}})")
        self._command("powerquery", "create", "--query-name", "AA",
                      "--load-destination", "connection-only", "--m-code", "#table(type table [X=number], {{2}})")
        snapshot = self._inspect("A1:B2")
        self.assertEqual(snapshot["sheets"][0]["sourceFormulas"][1][1], "=A2*3")
        self.assertIn(snapshot["sheets"][0]["sourceFormats"][1][1], ("0.00", "0,00"))
        self.assertEqual({query["name"] for query in snapshot["queries"]}, {"A", "AA"})

    def test_query_recovery_updates_the_surviving_query(self):
        with self.assertRaises(subprocess.CalledProcessError):
            self._command("powerquery", "create", "--query-name", "ReviewQuery", "--target-sheet", "QueryData",
                          "--m-code", 'error "Expected test load failure"')
        queries = self._command("powerquery", "list")
        self.assertEqual(len(queries["queries"]), 1)
        self.assertEqual(queries["queries"][0]["name"], "ReviewQuery")
        code = "#table(type table [X=number], {{1}})"
        self._command("powerquery", "evaluate", "--m-code", code)
        self._command("powerquery", "update", "--query-name", "ReviewQuery", "--m-code", code)
        result = self._command("range", "get-values", "--sheet", "QueryData", "--range", "A1:A2")
        self.assertEqual(result["values"], [["X"], [1]])

    def test_column_chart_values_labels_and_real_overlap(self):
        self._write([
            ["Month", "Revenue", "Expenses"],
            ["January", 50000, 35000], ["February", 55000, 38000],
            ["March", 48000, 32000], ["April", 62000, 41000], ["May", 58000, 39000],
        ], "A1:C6")
        self._command("chart", "create-from-range", "--sheet", "Sheet1", "--source-range-address", "A1:C6",
                      "--chart-type", "ColumnClustered", "--chart-name", "ReviewChart", "--target-range", "A8:H22")
        snapshot = self._inspect("A1:C6")
        assert_chart(snapshot, below=True)
        for defect in ("values", "categories", "type", "missing"):
            bad = copy.deepcopy(snapshot)
            chart = bad["sheets"][0]["charts"][0]
            if defect == "missing":
                chart["series"].pop()
            elif defect == "type":
                chart["type"] = 5
            else:
                chart["series"][0][defect][0] = "wrong"
            with self.subTest(defect=defect), self.assertRaises(AssertionError):
                assert_chart(bad, below=True)
        self._reopen()
        self._command("chart", "fit-to-range", "--sheet", "Sheet1", "--chart-name", "ReviewChart", "--range", "A1:H15")
        with self.assertRaises(AssertionError):
            assert_chart(self._inspect("A1:C6"), below=True)

    def test_line_chart_table_and_right_hand_position(self):
        self._write([
            ["Product", "Q1", "Q2", "Q3"], ["Widget", 100, 150, 120],
            ["Gadget", 80, 90, 110], ["Device", 200, 180, 220], ["Tool", 50, 60, 75],
        ], "A1:D5")
        self._command("table", "create", "--sheet", "Sheet1", "--range", "A1:D5", "--table-name", "ProductSales")
        self._command("chart", "create-from-table", "--sheet", "Sheet1", "--table-name", "ProductSales",
                      "--chart-type", "Line", "--chart-name", "Products", "--target-range", "F2:M16")
        assert_chart(self._inspect("A1:D5"), below=False)

    def test_table_slicer_reads_and_rejects_wrong_saved_selection(self):
        self._write([
            ["Department", "Employee", "Status", "Salary"],
            ["Engineering", "Alice", "Active", 85000], ["Engineering", "Bob", "Active", 92000],
            ["Marketing", "Carol", "Active", 78000], ["Marketing", "Dave", "Inactive", 70000],
            ["Sales", "Eve", "Active", 65000], ["Sales", "Frank", "Inactive", 62000],
            ["Engineering", "Grace", "Active", 88000], ["Sales", "Henry", "Active", 71000],
        ], "A1:D9")
        self._command("table", "create", "--sheet", "Sheet1", "--range", "A1:D9", "--table-name", "Employees")
        for field, position, selected in (("Department", "F2", "Engineering"), ("Status", "H2", "Active")):
            self._command("slicer", "create-table-slicer", "--table-name", "Employees", "--column-name", field,
                          "--slicer-name", f"{field}Slicer", "--destination-sheet", "Sheet1", "--position", position)
            self._command("slicer", "set-table-slicer-selection", "--slicer-name", f"{field}Slicer",
                          "--selected-items", json.dumps([selected]))
        assert_table_slicers(self._inspect())
        self._reopen()
        self._command("slicer", "set-table-slicer-selection", "--slicer-name", "DepartmentSlicer", "--selected-items", '["Marketing"]')
        with self.assertRaises(AssertionError):
            assert_table_slicers(self._inspect())

    def _pivot(self, table, sheet, row, value):
        self._command("sheet", "create", "--sheet", sheet)
        self._command("pivottable", "create-from-table", "--table-name", table, "--destination-sheet", sheet,
                      "--destination-cell", "A1", "--pivot-table-name", "SummaryPivot")
        self._command("pivottablefield", "add-row-field", "--pivot-table-name", "SummaryPivot", "--field-name", row)
        self._command("pivottablefield", "add-value-field", "--pivot-table-name", "SummaryPivot", "--field-name", value)
        self._command("pivottable", "refresh", "--pivot-table-name", "SummaryPivot")

    def test_pivot_slicer_has_actual_filtered_total(self):
        self._write([
            ["Region", "Product", "Quarter", "Sales"],
            ["North", "Laptop", "Q1", 15000], ["North", "Phone", "Q1", 8000],
            ["North", "Laptop", "Q2", 18000], ["North", "Phone", "Q2", 9500],
            ["South", "Laptop", "Q1", 12000], ["South", "Phone", "Q1", 7500],
            ["South", "Laptop", "Q2", 14000], ["South", "Phone", "Q2", 8200],
        ], "A1:D9")
        self._command("table", "create", "--sheet", "Sheet1", "--range", "A1:D9", "--table-name", "SalesData")
        self._pivot("SalesData", "Analysis", "Region", "Sales")
        self._command("slicer", "create-slicer", "--pivot-table-name", "SummaryPivot", "--field-name", "Region",
                      "--slicer-name", "RegionSlicer", "--destination-sheet", "Analysis", "--position", "E2")
        self._command("slicer", "set-slicer-selection", "--slicer-name", "RegionSlicer", "--selected-items", '["North"]')
        assert_pivot_slicer(self._inspect())
        self._reopen()
        self._command("slicer", "set-slicer-selection", "--slicer-name", "RegionSlicer", "--selected-items", '["South"]')
        with self.assertRaises(AssertionError):
            assert_pivot_slicer(self._inspect())

    def test_combined_slicers_filter_independent_objects(self):
        self._write([
            ["Category", "Product", "Warehouse", "Stock", "Price"],
            ["Electronics", "Laptop", "West", 50, 999], ["Electronics", "Phone", "West", 120, 599],
            ["Electronics", "Laptop", "East", 35, 999], ["Electronics", "Phone", "East", 80, 599],
            ["Furniture", "Desk", "West", 25, 350], ["Furniture", "Chair", "West", 40, 175],
            ["Furniture", "Desk", "East", 30, 350], ["Furniture", "Chair", "East", 55, 175],
        ], "A1:E9")
        self._command("table", "create", "--sheet", "Sheet1", "--range", "A1:E9", "--table-name", "Inventory")
        self._pivot("Inventory", "Summary", "Category", "Stock")
        self._command("slicer", "create-table-slicer", "--table-name", "Inventory", "--column-name", "Warehouse",
                      "--slicer-name", "WarehouseSlicer", "--destination-sheet", "Sheet1", "--position", "G2")
        self._command("slicer", "create-slicer", "--pivot-table-name", "SummaryPivot", "--field-name", "Category",
                      "--slicer-name", "CategorySlicer", "--destination-sheet", "Summary", "--position", "D2")
        self._command("slicer", "set-table-slicer-selection", "--slicer-name", "WarehouseSlicer", "--selected-items", '["West"]')
        self._command("slicer", "set-slicer-selection", "--slicer-name", "CategorySlicer", "--selected-items", '["Electronics"]')
        snapshot = self._inspect()
        assert_combined_slicers(snapshot)
        for sheet in snapshot["sheets"]:
            for pivot in sheet["pivots"]:
                pivot["values"][-1][-1] = 170
        with self.assertRaises(AssertionError):
            assert_combined_slicers(snapshot)


if __name__ == "__main__":
    unittest.main()
