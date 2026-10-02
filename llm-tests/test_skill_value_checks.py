"""Offline false-positive regressions for skill-value evidence."""

from __future__ import annotations

import unittest
import tempfile
import json
from pathlib import Path
from unittest.mock import AsyncMock

from pytest_skill_engineering.copilot.result import CopilotResult, ToolCall, Turn
from pytest_skill_engineering.copilot import DiscoveredSkill, SkillDiscovery

from skill_value import execution_metrics, assert_explicit_save, completed_operations, record_execution, summarize, paired_results
from skill_value_tasks import SkillTask, check_snapshot
from skill_discovery import assert_skill_exposure
from consent_scenarios import assert_read_only
from cli_evidence import batch_steps, tool_output


class SkillValueChecks(unittest.TestCase):
    def test_control_rejects_unformatted_monetary_totals(self):
        data = [["Month", "Revenue", "Expenses"], ["January", 135, 82],
                ["February", 210, 119], ["March", 178, 96]]
        values = [*data, [None, None, None], [None, 523, 297]]
        formulas = [*data, [None, None, None], [None, "=SUM(B2:B4)", "=SUM(C2:C4)"]]
        formats = [["General"] * 3, *[["General", "0.00", "0.00"] for _ in range(3)],
                   ["General"] * 3, ["General"] * 3]
        snapshot = {"sheets": [{"name": "Sheet1", "sourceValues": values, "sourceFormulas": formulas,
                               "sourceFormats": formats, "tables": [{"name": "SalesReport", "rows": data[1:]}]}]}
        with self.assertRaises(AssertionError):
            check_snapshot("control", snapshot, None)

    def test_successful_screenshot_metadata_is_not_parsed_as_service_json(self):
        call = ToolCall(
            "excel-mcp-screenshot", {"action": "capture"},
            result='Screenshot of Sheet1.\n\n{"success":true,"sheetName":"Sheet1"}',
            completion_received=True, success=True,
        )
        result = CopilotResult(turns=[Turn("assistant", "", [call])])
        self.assertEqual(completed_operations(result), ["screenshot.capture"])

    def test_discovery_rejects_same_name_from_a_different_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            wrong = str(Path(directory) / "unexpected" / "SKILL.md")
            with self.assertRaisesRegex(AssertionError, "Wrong skill source"):
                assert_skill_exposure([{"name": "excel-cli", "enabled": True, "path": wrong}],
                                      "excel-cli", directory)

    def test_mcp_text_and_structured_duplicates_are_one_operation(self):
        payload = {"message": "Calculation complete for range 'D2:D121'", "success": True}
        call = ToolCall("excel-mcp-calculation_mode", {"action": "calculate"}, result=(
            json.dumps(payload) + "\n\n" + json.dumps(payload)
        ), completion_received=True, success=True)
        result = CopilotResult(turns=[Turn("assistant", "", [call])])
        self.assertEqual(completed_operations(result), ["calculation_mode.calculate"])

    def test_cli_sdk_text_and_structured_duplicates_must_agree(self):
        payload = {"exit_code": 0, "stdout": '{"success":true}'}
        for duplicate, valid in ((payload, True), ({**payload, "exit_code": 1}, False)):
            call = ToolCall("excel_execute", {}, result=(
                json.dumps(payload) + "\n\n" + json.dumps({"result": json.dumps(duplicate)})
            ))
            if valid:
                self.assertEqual(tool_output(call), payload)
            else:
                with self.assertRaises(AssertionError):
                    tool_output(call)

    def test_control_preparation_needs_no_existing_workbook_or_excel_process(self):
        with tempfile.TemporaryDirectory() as directory:
            task = SkillTask(Path(directory) / "workbook.xlsx", "test", Path("not-an-executable"))
            task.prepare("control")
            self.assertFalse(task.path.exists())
            self.assertIsNone(task.before)

    def test_save_claim_and_failed_close_are_not_persistence(self):
        for calls in ([], [ToolCall(
            "excel-mcp-file", {"action": "close", "save": True},
            result='{"success":false}', completion_received=True, success=True,
        )]):
            with self.subTest(calls=calls), self.assertRaises(AssertionError):
                assert_explicit_save(CopilotResult(turns=[Turn("assistant", "Saved.", calls)]), "mcp")

    def test_successful_batch_close_uses_the_executed_save_argument(self):
        for saving in (True, False):
            commands = [{"command": "session.close", "args": {"save": saving}}]
            calls = [
                ToolCall("excel_execute", {"args": "batch --help"},
                         result='{"exit_code":0,"stdout":"Batch command help"}',
                         completion_received=True, success=True),
                ToolCall("workspace", {"action": "write", "path": "commands.json",
                                      "content": json.dumps(commands)},
                         result="Written", completion_received=True, success=True),
                ToolCall("excel_execute", {"args": "-q batch --input commands.json"},
                         result=json.dumps({"exit_code": 0, "stdout": json.dumps({
                             "index": 0, "command": "session.close", "success": True, "result": None,
                         })}), completion_received=True, success=True),
            ]
            result = CopilotResult(turns=[Turn("assistant", "", calls)])
            if saving:
                assert_explicit_save(result, "cli")
            else:
                with self.assertRaises(AssertionError):
                    assert_explicit_save(result, "cli")

    def test_rejected_batch_syntax_is_not_an_executed_batch(self):
        calls = [
            ToolCall("excel_execute", {"args": "batch --commands '[]'"},
                     result=json.dumps({"exit_code": 1, "stdout": json.dumps({
                         "success": False, "exceptionType": "CommandParseException",
                         "error": "Unknown option 'commands'.",
                     })}), completion_received=True, success=True),
            ToolCall("excel_execute", {"args": "session close --session owned --save"},
                     result='{"exit_code":0,"stdout":"{\\"success\\":true}"}',
                     completion_received=True, success=True),
        ]
        result = CopilotResult(turns=[Turn("assistant", "", calls)])
        assert_explicit_save(result, "cli")
        self.assertEqual(completed_operations(result), ["session.close"])

    def test_indexed_failed_batch_step_is_not_a_zero_execution_parser_rejection(self):
        command = {"command": "range.set-values", "args": {"values": [[1]]}}
        calls = [
            ToolCall("workspace", {"action": "write", "path": "commands.json",
                                  "content": json.dumps([command])},
                     completion_received=True, success=True),
            ToolCall("excel_execute", {"args": "batch --input commands.json"},
                     result=json.dumps({"exit_code": 1, "stdout": json.dumps({
                         "index": 0, "command": command["command"], "success": False,
                         "exceptionType": "CommandParseException",
                     })}), completion_received=True, success=True),
        ]
        result = CopilotResult(turns=[Turn("assistant", "", calls)])
        steps = batch_steps(result, calls[-1])
        self.assertEqual(len(steps), 1)
        self.assertFalse(steps[0]["success"])
        self.assertEqual(steps[0]["args"], command["args"])
        with self.assertRaises(AssertionError):
            assert_read_only(result, require_read=False)

    def test_parser_rejection_requires_a_known_nonzero_exit_code(self):
        for code in (None, 0):
            call = ToolCall(
                "excel_execute", {"args": "batch --commands '[]'"},
                result=json.dumps({"exit_code": code, "stdout": json.dumps({
                    "success": False, "exceptionType": "CommandParseException",
                })}), completion_received=True, success=True,
            )
            with self.subTest(code=code), self.assertRaises(AssertionError):
                batch_steps(CopilotResult(turns=[Turn("assistant", "", [call])]), call)

    def test_batch_read_only_checks_reject_embedded_mutations(self):
        for command in ("range.get-values", "range.set-values"):
            commands = [{"command": command, "args": {}}]
            calls = [
                ToolCall("workspace", {"action": "write", "path": "commands.json",
                                      "content": json.dumps(commands)},
                         result="Written", completion_received=True, success=True),
                ToolCall("excel_execute", {"args": "-q batch --input commands.json"},
                         result=json.dumps({"exit_code": 0, "stdout": json.dumps({
                             "index": 0, "command": command, "success": True,
                         })}), completion_received=True, success=True),
            ]
            result = CopilotResult(turns=[Turn("assistant", "", calls)])
            if command == "range.get-values":
                assert_read_only(result)
            else:
                with self.assertRaises(AssertionError):
                    assert_read_only(result)

    def test_missing_usage_is_unknown_not_zero(self):
        metrics = execution_metrics(CopilotResult())
        self.assertIsNone(metrics["tokens"])
        self.assertIsNone(metrics["premium_requests"])

    def test_direct_skill_reads_are_recorded_but_task_reads_are_not(self):
        calls = [ToolCall("workspace", {"action": "read", "path": path})
                 for path in ("SKILL.md", r"skill\references\range.md", "price-updates.json")]
        metrics = execution_metrics(CopilotResult(turns=[Turn("assistant", "", calls)]))
        self.assertEqual(len(metrics["skill_reads"]), 2)

    def test_missing_disabled_and_ambient_skills_reject_comparisons(self):
        for discovered, expected in (
            ([], "excel-mcp"),
            ([{"name": "excel-mcp", "enabled": False}], "excel-mcp"),
            ([{"name": "ambient", "enabled": True}], None),
        ):
            with self.subTest(discovered=discovered), self.assertRaises(AssertionError):
                assert_skill_exposure(discovered, expected)
        assert_skill_exposure([], None)
        assert_skill_exposure([{"name": "excel-cli", "enabled": True}], "excel-cli")

    def test_comparison_uses_individual_verification_not_shared_outcome(self):
        baseline = {"transport": "mcp", "condition": "without-skill", "task": "control",
                    "passed": True, "category": "verified", "tokens": 100,
                    "premium_requests": None, "skill_reads": [], "tool_calls": 2}
        treatment = {**baseline, "condition": "with-skill", "passed": False,
                     "category": "task", "tokens": 20}
        for records in ([baseline, treatment], [
            {**baseline, "passed": False}, {**treatment, "passed": True},
        ]):
            result = summarize(records)
            self.assertEqual(
                {row["condition"]: row["verified"] for row in result},
                {record["condition"]: int(record["passed"]) for record in records},
            )

    def test_failed_attempt_usage_counts_in_cost_per_completed_task(self):
        records = [{"transport": "cli", "condition": "with-skill", "task": "control",
                    "passed": True, "category": "verified", "tokens": 100, "premium_requests": 1},
                   {"transport": "cli", "condition": "with-skill", "task": "control",
                    "passed": False, "category": "task", "tokens": 200, "premium_requests": 2}]
        row = summarize(records)[0]
        self.assertEqual(row["tokens_per_verified"], 300)
        self.assertEqual(row["requests_per_verified"], 3)
        records[1]["tokens"] = None
        self.assertIsNone(summarize(records)[0]["tokens_per_verified"])

    def test_paired_report_keeps_missing_and_failed_conditions_visible(self):
        baseline = {"transport": "mcp", "task": "control", "repetition": 1,
                    "condition": "without-skill", "passed": True, "tokens": 100, "skill_reads": []}
        treatment = {**baseline, "condition": "with-skill", "passed": False, "tokens": 80}
        row = paired_results([baseline, treatment])[0]
        self.assertFalse(row["both_verified"])
        self.assertIsNone(row["token_saving_percent"])
        treatment["passed"] = True
        self.assertEqual(paired_results([baseline, treatment])[0]["token_saving_percent"], 20)
        treatment["tokens"] = None
        self.assertIsNone(paired_results([baseline, treatment])[0]["token_saving_percent"])
        self.assertIsNone(paired_results([baseline])[0]["with_skill"])


class ExecutionEvidenceChecks(unittest.IsolatedAsyncioTestCase):
    async def test_selection_result_is_saved_separately_from_workbook_correctness(self):
        for expected_read in (True, False):
            with tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "attempt.json"
                record = {"passed": False, "condition": "without-skill", "transport": "mcp",
                          "expected_skill_read": expected_read}
                result = CopilotResult(success=True, stop_reason="completed")
                await record_execution(record, path, AsyncMock(return_value=result), None,
                                       "request", lambda result: {"correct_workbook": True})
                saved = json.loads(path.read_text(encoding="utf-8"))
                self.assertTrue(saved["passed"])
                self.assertEqual(saved["selection_matches_intent"], not expected_read)

    async def test_progress_survives_an_abrupt_evaluation_interruption(self):
        from types import SimpleNamespace
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "attempt.json"
            record = {"passed": False, "condition": "without-skill", "transport": "mcp"}
            agent = SimpleNamespace(extra_config={})
            event = SimpleNamespace(model_dump_json=lambda: '{"type":"assistant.message","data":{"content":"working"}}')

            async def evaluate(agent, prompt):
                agent.extra_config["on_event"](event)
                saved = json.loads(path.read_text(encoding="utf-8"))
                self.assertEqual(saved["status"], "running")
                self.assertEqual(saved["events_received"], 1)
                raise KeyboardInterrupt("Simulated process interruption")

            with self.assertRaises(KeyboardInterrupt):
                await record_execution(record, path, evaluate, agent, "request", lambda result: {})
            saved = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(saved["status"], "interrupted")
            self.assertIsNone(saved.get("tokens"))
            journal = path.with_suffix(".events.jsonl")
            self.assertEqual(json.loads(journal.read_text(encoding="utf-8"))["type"], "assistant.message")
            self.assertNotIn("on_event", agent.extra_config)

    async def test_returned_execution_is_saved_before_verification(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "attempt.json"
            record = {"passed": False, "condition": "without-skill", "transport": "mcp"}
            result = CopilotResult(success=True, stop_reason="completed")

            def verify(result):
                saved = json.loads(path.read_text(encoding="utf-8"))
                self.assertEqual(saved["status"], "verifying")
                self.assertTrue(saved["execution"]["success"])
                return {}

            await record_execution(record, path, AsyncMock(return_value=result), None, "request", verify)
            self.assertEqual(json.loads(path.read_text(encoding="utf-8"))["status"], "finished")

    async def test_treatment_requires_actual_session_discovery_evidence(self):
        with tempfile.TemporaryDirectory() as directory:
            record = {"passed": False, "condition": "with-skill", "transport": "mcp"}
            result = CopilotResult(success=True, stop_reason="completed")
            await record_execution(record, Path(directory) / "attempt.json",
                                   AsyncMock(return_value=result), None, "request", lambda result: {})
            self.assertEqual(record["category"], "capture")
            self.assertFalse(record["passed"])

    async def test_treatment_records_matching_availability_separately_from_reads(self):
        discovery = SkillDiscovery(
            skills=[DiscoveredSkill(name="excel-mcp", description="Test", source="explicit",
                                    enabled=True, user_invocable=True)],
            complete=True,
        )
        with tempfile.TemporaryDirectory() as directory:
            record = {"passed": False, "condition": "with-skill", "transport": "mcp"}
            result = CopilotResult(success=True, stop_reason="completed", skill_discovery=discovery)
            path = Path(directory) / "attempt.json"
            await record_execution(record, path, AsyncMock(return_value=result), None, "request", lambda result: {})
            self.assertTrue(record["passed"])
            self.assertEqual(record["skill_reads"], [])
            self.assertTrue(json.loads(path.read_text())["skill_discovery"]["complete"])

    async def test_treatment_accepts_the_declared_scoped_skill(self):
        discovery = SkillDiscovery(
            skills=[DiscoveredSkill(name="excel-mcp-report-formatting", description="Test", source="explicit",
                                    enabled=True, user_invocable=True)],
            complete=True,
        )
        with tempfile.TemporaryDirectory() as directory:
            record = {"passed": False, "condition": "with-skill", "transport": "mcp",
                      "skill_name": "excel-mcp-report-formatting"}
            result = CopilotResult(success=True, stop_reason="completed", skill_discovery=discovery)
            await record_execution(record, Path(directory) / "attempt.json",
                                   AsyncMock(return_value=result), None, "request", lambda result: {})
            self.assertTrue(record["passed"])

    async def test_treatment_rejects_disabled_or_different_skill(self):
        for name, enabled in (("excel-mcp", False), ("excel-cli", True)):
            with self.subTest(name=name, enabled=enabled), tempfile.TemporaryDirectory() as directory:
                discovery = SkillDiscovery(
                    skills=[DiscoveredSkill(name=name, description="Test", source="explicit",
                                            enabled=enabled, user_invocable=True)],
                    complete=True,
                )
                record = {"passed": False, "condition": "with-skill", "transport": "mcp"}
                result = CopilotResult(success=True, stop_reason="completed", skill_discovery=discovery)
                await record_execution(record, Path(directory) / "attempt.json",
                                       AsyncMock(return_value=result), None, "request", lambda result: {})
                self.assertFalse(record["passed"])
                self.assertEqual(record["category"], "capture")

    def test_unchecked_baseline_is_not_inferred_empty_discovery(self):
        self.assertIsNone(execution_metrics(CopilotResult())["skill_discovery"])

    async def test_timeout_keeps_its_budget_classification_with_incomplete_calls(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "attempt.json"
            record = {"passed": False}
            result = CopilotResult(
                success=False, stop_reason="timeout", error="Timed out",
                turns=[Turn("assistant", "", [ToolCall("excel-mcp-file", {"action": "open"})])],
            )
            await record_execution(record, path, AsyncMock(return_value=result), None, "request", lambda result: {})
            self.assertEqual(record["category"], "budget")
            self.assertFalse(record["evidence_complete"])
            self.assertFalse(record["passed"])

    async def test_execution_exception_is_persisted_and_rethrown(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "attempt.json"
            record = {"passed": False, "category": "execution"}
            evaluate = AsyncMock(side_effect=RuntimeError("Runtime failed"))
            with self.assertRaisesRegex(RuntimeError, "Runtime failed"):
                await record_execution(record, path, evaluate, None, "request", lambda result: {})
            saved = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(saved["category"], "harness")
            self.assertIn("Runtime failed", saved["verification_error"])
            self.assertFalse(saved["passed"])

    async def test_checker_exception_preserves_returned_execution(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "attempt.json"
            record = {"passed": False, "category": "execution"}
            result = CopilotResult(success=True, model_used="gpt-6.1-sol", stop_reason="completed")

            def verify(result):
                raise OSError("Workbook inspection failed")

            with self.assertRaisesRegex(OSError, "Workbook inspection failed"):
                await record_execution(record, path, AsyncMock(return_value=result), None, "request", verify)
            saved = json.loads(path.read_text(encoding="utf-8"))
            self.assertTrue(saved["execution"]["success"])
            self.assertEqual(saved["category"], "harness")
            self.assertIn("Workbook inspection failed", saved["verification_error"])

    async def test_wrong_workbook_is_a_task_failure_not_a_harness_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "attempt.json"
            record = {"passed": False, "category": "execution"}
            result = CopilotResult(success=True, stop_reason="completed")

            def verify(result):
                raise AssertionError("Wrong total")

            await record_execution(record, path, AsyncMock(return_value=result), None, "request", verify)
            saved = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(saved["category"], "task")
            self.assertEqual(saved["verification_error"], "Wrong total")

    async def test_missing_batch_input_is_capture_failure_not_agent_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "attempt.json"
            record = {"passed": False, "condition": "without-skill", "transport": "cli"}
            call = ToolCall(
                "excel_execute", {"args": "batch --input missing.json"},
                result=json.dumps({"exit_code": 0, "stdout": json.dumps({
                    "index": 0, "command": "session.close", "success": True,
                })}), completion_received=True, success=True,
            )
            result = CopilotResult(success=True, stop_reason="completed",
                                   turns=[Turn("assistant", "", [call])])
            await record_execution(record, path, AsyncMock(return_value=result), None,
                                   "request", completed_operations)
            saved = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(saved["category"], "capture")
            self.assertFalse(saved["passed"])
            self.assertTrue(saved["execution"]["success"])
            self.assertIn("captured workspace write", saved["verification_error"])


if __name__ == "__main__":
    unittest.main()
