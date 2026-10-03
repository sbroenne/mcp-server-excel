"""Check consent assertions without Excel or an external model."""

from __future__ import annotations

import copy
import unittest
from pathlib import Path
from unittest.mock import patch

from pytest_skill_engineering.copilot.result import CopilotResult, ToolCall, Turn

from consent_scenarios import ConsentWorkbook, _BASE_VALUES, assert_consent_outcome, assert_read_only


class ConsentOutcomeTests(unittest.TestCase):
    def test_canonical_calculation_reads_are_allowed_for_both_entry_points(self):
        calls = [
            ToolCall("excel_execute", {"args": "calculationmode get-settings --session test"}),
            ToolCall("excel-mcp-calculation_mode", {"action": "get-settings", "session_id": "test"}),
        ]
        for call in calls:
            with self.subTest(call=call):
                assert_read_only(CopilotResult(turns=[Turn("assistant", "", [call])]), require_read=False)

    def test_calculation_settings_writes_are_not_read_only(self):
        call = ToolCall("excel-mcp-calculation_mode", {"action": "set-settings", "session_id": "test", "mode": "manual"})
        with self.assertRaisesRegex(AssertionError, "State-changing call"):
            assert_read_only(CopilotResult(turns=[Turn("assistant", "", [call])]), require_read=False)

    def setUp(self):
        self.workbook = ConsentWorkbook(Path("budget.xlsx"), "test", Path("excelcli.exe"))
        values = copy.deepcopy(_BASE_VALUES)
        values[2][2] = 480
        values[4][2] = 1480
        snapshot = {"sheets": [{"tables": [], "charts": [], "pivots": [], "sourceValues": values}]}
        self.sessions = []
        cleanup = patch.object(self.workbook, "close_owned_sessions")
        self.cleanup = cleanup.start()
        self.addCleanup(cleanup.stop)
        cli = patch.object(self.workbook, "cli", side_effect=self.cli_response)
        cli.start()
        self.addCleanup(cli.stop)
        inspect = patch("consent_scenarios.read_saved_workbook", return_value=snapshot)
        inspect.start()
        self.addCleanup(inspect.stop)

    def cli_response(self, *args):
        if args[:2] == ("session", "list"):
            return {"sessions": self.sessions}
        if args[:2] == ("session", "open"):
            return {"sessionId": "inspection"}
        if args[:2] == ("range", "get-formulas"):
            return {"formulas": [["=SUM(C2:C3)"]]}
        return {"success": True}

    def assert_outcome(self, calls):
        for call in calls:
            if call.result is None:
                call.result = ('{"exit_code":0,"stdout":"{\\"success\\":true}"}'
                               if "execute" in call.name else '{"success":true}')
                call.completion_received = True
                call.success = True
        result = CopilotResult(turns=[Turn("assistant", "Changed Food to 480.", calls)])
        assert_consent_outcome(result, self.workbook, "clear-edit", [])

    def test_rejects_missing_or_non_saving_close_before_cleanup(self):
        cases = [
            [ToolCall("excel_execute", {"args": "range set-values --session test"})],
            [ToolCall("excel_execute", {"args": "session close --session test"})],
            [ToolCall("excel_execute", {"args": "session close --session test --save false"})],
            [ToolCall("excel_execute", {"args": "session close --session test --save=false"})],
            [ToolCall("excel-mcp-range", {"action": "set-values", "session_id": "test"})],
            [ToolCall("excel-mcp-file", {"action": "close", "session_id": "test", "save": False})],
            [ToolCall("excel-mcp-file", {"action": "close", "session_id": "test", "save": "true"})],
            [
                ToolCall("excel-mcp-file", {"action": "close", "session_id": "test", "save": True}),
                ToolCall("excel-mcp-file", {"action": "open", "path": "budget.xlsx"}),
            ],
        ]
        for calls in cases:
            with self.subTest(calls=calls):
                with self.assertRaisesRegex(AssertionError, "save and close"):
                    self.assert_outcome(calls)
                self.cleanup.assert_not_called()

    def test_rejects_a_live_cli_session_before_cleanup(self):
        self.sessions = [{"filePath": str(self.workbook.path), "sessionId": "test"}]
        with self.assertRaisesRegex(AssertionError, "still open"):
            self.assert_outcome([
                ToolCall("excel_execute", {"args": "session close --session test --save"}),
            ])
        self.cleanup.assert_not_called()

    def test_accepts_explicit_saving_close_for_both_entry_points(self):
        calls = [
            ToolCall("excel_execute", {"args": "session close --session test --save"}),
            ToolCall("excel_execute", {"args": "session close --session test --save true"}),
            ToolCall("excel_execute", {"args": "session close --session test --save=true"}),
            ToolCall("excel-mcp-file", {"action": "close", "session_id": "test", "save": True}),
        ]
        for call in calls:
            with self.subTest(call=call):
                self.assert_outcome([call])

    def test_rejects_a_completed_but_failed_save(self):
        with self.assertRaises(AssertionError):
            self.assert_outcome([ToolCall(
                "excel-mcp-file", {"action": "close", "session_id": "test", "save": True},
                result='{"success":false}', completion_received=True, success=True,
            )])


if __name__ == "__main__":
    unittest.main()
