"""Real Excel proofs that synthetic fixtures and outcome checks work without AI."""

from __future__ import annotations

import copy
import json
import tempfile
import unittest
import uuid
from pathlib import Path

from skill_value_tasks import SkillTask, NEW_ORDERS, check_snapshot
from workbook_assertions import read_saved_workbook
from pytest_skill_engineering.copilot.result import CopilotResult, ToolCall, Turn


class SkillValuePreparation(unittest.TestCase):
    def test_control(self):
        self._check_tasks(("control",))

    def test_bulk_update(self):
        self._check_tasks(("bulk-update",))

    def test_query_recovery(self):
        self._check_tasks(("query-recovery",))

    def test_model_refresh(self):
        self._check_tasks(("model-refresh",))

    def test_read_only_audit(self):
        self._check_tasks(("read-only-audit",))

    def test_formatting_reports(self):
        root = Path(__file__).resolve().parents[1]
        with tempfile.TemporaryDirectory(prefix="formatting-check-") as directory:
            task = SkillTask(Path(directory) / "workbook.xlsx", f"formatting-check-{uuid.uuid4().hex}",
                             root / "src" / "ExcelMcp.CLI" / "bin" / "Release" / "net10.0-windows" / "excelcli.exe")
            try:
                task.prepare("financial-formatting")
                with self.assertRaises(AssertionError):
                    check_snapshot("financial-formatting", task.before, task.before)
                task.session = task.cli("session", "open", str(task.path))["sessionId"]
                task.command("rangeformat", "format", "--sheet", "Sheet1", "--range-addresses", "A1:C1",
                             "--format-options", '{"bold":true}')
                task.command("rangeformat", "format", "--sheet", "Sheet1", "--range-addresses", "B2:C4",
                             "--format-options", '{"fontColor":"#0000FF"}')
                task.command("rangeformat", "format", "--sheet", "Sheet1", "--range-addresses", "B5:C5",
                             "--format-options", '{"fontColor":"#000000"}')
                task.command("range", "set-number-format", "--sheet", "Sheet1", "--range", "B2:B5",
                             "--format-code", '$#,##0.00;($#,##0.00);"-"')
                task.command("range", "set-number-format", "--sheet", "Sheet1", "--range", "C2:C5", "--format-code", "0.0%")
                task.command("rangeformat", "auto-fit-columns", "--sheet", "Sheet1", "--range", "A:C")
                task.command("session", "close", "--save")
                task.session = None
                after = read_saved_workbook(str(task.path), task.source_range, include_presentation=True)
                for name in ("formatting-report", "financial-formatting"):
                    check_snapshot(name, after, task.before)
                invalid = copy.deepcopy(after)
                invalid["sheets"][0]["sourceValues"][1][2] = 45
                with self.assertRaises(AssertionError):
                    check_snapshot("formatting-report", invalid, task.before)
            finally:
                try:
                    task.close_owned_sessions()
                finally:
                    task.cli("service", "stop")

    def _check_tasks(self, names):
        for name in names:
            with self.subTest(task=name), tempfile.TemporaryDirectory(prefix="skill-check-") as directory:
                root = Path(__file__).resolve().parents[1]
                task = SkillTask(Path(directory) / "workbook.xlsx", f"skill-check-{uuid.uuid4().hex}",
                                 root / "src" / "ExcelMcp.CLI" / "bin" / "Release" / "net10.0-windows" / "excelcli.exe")
                try:
                    task.prepare(name)
                    if name == "read-only-audit":
                        self.assertEqual(task.before["sheets"][0]["sourceValues"][4][2], 1450)
                        self.assertEqual(task.before["sheets"][0]["sourceFormulas"][4][2], "=SUM(C2:C3)")
                        task.session = task.cli("session", "open", str(task.path))["sessionId"]
                        task.command("calculationmode", "set-settings", "--mode", "manual")
                        task.values("Sheet1", "B7", [["Unsaved user note"]])
                        values = task.command("range", "get-values", "--sheet", "Sheet1", "--range", "A1:C8")
                        call = ToolCall(
                            "excel_execute", {"args": "range get-values"},
                            result=json.dumps({"exit_code": 0, "stdout": json.dumps(values)}),
                            completion_received=True, success=True,
                        )
                        result = CopilotResult(
                            success=True, model_used="gpt-6.1-sol", stop_reason="completed",
                            turns=[Turn("assistant", "The actual total formula is incorrect; the correct total is 1630.", [call])],
                        )
                        task.verify(name, result, "cli")
                        task.values("Sheet1", "B7", [["Lost the user's note"]])
                        with self.assertRaises(AssertionError):
                            task.verify(name, result, "cli")
                        continue
                    if name != "control":
                        with self.assertRaises(AssertionError):
                            check_snapshot(name, task.before, task.before)
                        task.session = task.cli("session", "open", str(task.path))["sessionId"]
                    else:
                        task.session = task.cli("session", "create", str(task.path))["sessionId"]
                    if name == "control":
                        task.values("Sheet1", "A1:C4", [["Month", "Revenue", "Expenses"],
                                                       ["January", 135, 82], ["February", 210, 119], ["March", 178, 96]])
                        task.command("table", "create", "--sheet", "Sheet1", "--range", "A1:C4", "--table-name", "SalesReport")
                        task.command("range", "set-formulas", "--sheet", "Sheet1", "--range", "B6:C6",
                                     "--formulas", '[["=SUM(B2:B4)","=SUM(C2:C4)"]]')
                        task.command("range", "set-number-format", "--sheet", "Sheet1", "--range", "B2:C4",
                                     "--format-code", "0.00")
                        task.command("range", "set-number-format", "--sheet", "Sheet1", "--range", "B6:C6",
                                     "--format-code", "0.00")
                    elif name == "bulk-update":
                        task.values("Sheet1", "B2:B61", [[100 + i * 2] for i in range(1, 61)])
                        task.command("calculationmode", "calculate", "--scope", "application")
                    elif name == "query-recovery":
                        path = str(task.path.parent / "orders.csv").replace('"', '""')
                        code = (
                            f'let S = Csv.Document(File.Contents("{path}"),[Delimiter=",",Encoding=65001]), '
                            'H = Table.PromoteHeaders(S), '
                            'T = Table.TransformColumnTypes(H,{{"OrderDate",type date},{"Quantity",Int64.Type},{"Amount",type number}}) in T'
                        )
                        task.command("powerquery", "evaluate", "--m-code", code)
                        task.command("powerquery", "update", "--query-name", "ImportA", "--m-code", code)
                        task.command("powerquery", "load-to", "--query-name", "ImportA", "--load-destination", "worksheet",
                                     "--target-sheet", "Imported", "--target-cell-address", "B4")
                    else:
                        task.command("table", "append", "--table-name", "Orders", "--rows", json.dumps(NEW_ORDERS))
                        task.command("datamodel", "refresh")
                        task.command("pivottable", "refresh", "--pivot-table-name", "RevenuePivot")
                    task.command("session", "close", "--save")
                    task.session = None
                    after = read_saved_workbook(str(task.path), task.source_range, include_analysis=name == "model-refresh")
                    check_snapshot(name, after, task.before)
                    if name == "bulk-update":
                        wrong_chart = copy.deepcopy(after)
                        sheet = next(sheet for sheet in wrong_chart["sheets"] if sheet["name"] == "Sheet1")
                        sheet["charts"][0]["series"][0]["values"][0] = -1
                        with self.assertRaises(AssertionError):
                            check_snapshot(name, wrong_chart, task.before)
                    wrong = copy.deepcopy(after)
                    if name == "control":
                        wrong["sheets"][0]["sourceFormulas"][5][1] = 523
                    elif name == "bulk-update":
                        wrong["sheets"][0]["sourceFormulas"][1][3] = 408
                    elif name == "query-recovery":
                        wrong["queries"].append({"name": "ImportA_copy", "formula": "wrong"})
                    else:
                        next(sheet for sheet in wrong["sheets"] if sheet["name"] == "Analysis")["charts"][0]["pivot"] = None
                    with self.assertRaises(AssertionError):
                        check_snapshot(name, wrong, task.before)
                finally:
                    try:
                        task.close_owned_sessions()
                    finally:
                        task.cli("service", "stop")


if __name__ == "__main__":
    unittest.main()
