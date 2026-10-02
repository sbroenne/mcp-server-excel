"""Real Excel fixture/checker proofs, with no agent or model calls."""

from __future__ import annotations

import copy
import hashlib
import json
import tempfile
import unittest
import uuid
from pathlib import Path

from pytest_skill_engineering.copilot.result import CopilotResult, ToolCall, Turn

from spreadsheetbench import CASES, SpreadsheetBenchTask, check_snapshot, region
from workbook_assertions import _one


class SpreadsheetBenchWorkbookChecks(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix="excel-external-case-")
        self.addCleanup(self.directory.cleanup)
        self.task = SpreadsheetBenchTask(
            Path(self.directory.name) / "workbook.xlsx", f"external-proof-{uuid.uuid4().hex}",
            Path(__file__).resolve().parents[1] / "src" / "ExcelMcp.CLI" / "bin" / "Release"
            / "net10.0-windows" / "excelcli.exe",
        )
        self.addCleanup(self._close)

    def _close(self):
        try:
            self.task.close_owned_sessions()
        finally:
            self.task.cli("service", "stop")

    def test_each_official_answer_region_passes_but_original_and_wrong_answers_fail(self):
        for id, spec in CASES.items():
            with self.subTest(id=id):
                self.task.prepare(f"spreadsheetbench-{id}")
                self.assertEqual(list(Path(self.directory.name).iterdir()), [self.task.path])
                self.assertEqual(self.task.path.read_bytes(), self.task.loaded.input_bytes)
                self.assertEqual(hashlib.sha256(self.task.path.read_bytes()).hexdigest(), self.task.original_hash)
                check_snapshot(spec, self.task.expected, self.task.before, self.task.expected)
                with self.assertRaisesRegex(AssertionError, "Answer"):
                    check_snapshot(spec, self.task.before, self.task.before, self.task.expected)
                wrong = copy.deepcopy(self.task.expected)
                sheet = _one(wrong["sheets"], name=spec.sheet)
                row, column = min(region(spec.address))
                sheet["sourceValues"][row - sheet["sourceRow"]][column - sheet["sourceColumn"]] = "Wrong"
                with self.assertRaisesRegex(AssertionError, "Answer"):
                    check_snapshot(spec, wrong, self.task.before, self.task.expected)

    def test_saved_formula_answer_responds_to_new_data_but_constant_formulas_fail(self):
        task = self.task
        task.prepare("spreadsheetbench-38823")
        task.path.write_bytes(task.loaded.answer_bytes)
        result = CopilotResult(
            success=True, model_used="gpt-6.1-sol",
            turns=[Turn("assistant", "", [ToolCall(
                "excel-mcp-file", {"action": "close", "save": True},
                result='{"success":true}', completion_received=True, success=True,
            )])],
        )
        saved_hash = hashlib.sha256(task.path.read_bytes()).hexdigest()
        verified = task.verify("spreadsheetbench-38823", result, "mcp")
        self.assertTrue(verified["formula_probe_passed"])
        self.assertEqual(hashlib.sha256(task.path.read_bytes()).hexdigest(), saved_hash)
        task.session = task.cli("session", "open", str(task.path))["sessionId"]
        task.command(
            "range", "set-formulas", "--sheet", "Sheet1", "--range", "I4:I7",
            "--formulas", json.dumps([["=22"], ["=12"], ["=9"], ["=4"]]),
        )
        task.command("session", "close", "--save")
        task.session = None
        wrong_hash = hashlib.sha256(task.path.read_bytes()).hexdigest()
        with self.assertRaisesRegex(AssertionError, "Answer"):
            task.verify("spreadsheetbench-38823", result, "mcp")
        self.assertEqual(hashlib.sha256(task.path.read_bytes()).hexdigest(), wrong_hash)


if __name__ == "__main__":
    unittest.main()
