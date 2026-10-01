"""Check CLI call recording without Excel or an external model."""

from __future__ import annotations

import json
import unittest

from pytest_skill_engineering.copilot.result import CopilotResult, ToolCall, Turn

from conftest import _parse_cli_results, assert_cli_args_contain, assert_cli_exit_codes


def recorded(*calls: ToolCall) -> CopilotResult:
    return CopilotResult(turns=[Turn("assistant", "", list(calls))])


def execution(name: str, exit_code: int = 0, args: str = "chart read") -> ToolCall:
    return ToolCall(
        name, {"args": args},
        json.dumps({"exit_code": exit_code, "stdout": "", "stderr": ""}),
    )


class CliResultAssertionTests(unittest.TestCase):
    def test_accepts_both_exact_cli_tool_names(self):
        for name in ("excel_execute", "excel-cli-excel_execute"):
            with self.subTest(name=name):
                result = recorded(execution(name))
                assert_cli_exit_codes(result)
                assert_cli_exit_codes(result, strict=True)

    def test_checks_arguments_for_both_exact_cli_tool_names(self):
        for name in ("excel_execute", "excel-cli-excel_execute"):
            with self.subTest(name=name):
                result = recorded(execution(name, args="range set-values --values-file data.json"))
                assert_cli_args_contain(result, "--values-file")
                with self.assertRaisesRegex(AssertionError, "none did"):
                    assert_cli_args_contain(result, "--formulas-file")

    def test_preserves_interleaved_call_order_across_turns(self):
        result = CopilotResult(turns=[
            Turn("assistant", "", [execution("excel-cli-excel_execute", 1)]),
            Turn("assistant", "", [execution("excel_execute", 0)]),
            Turn("assistant", "", [execution("excel-cli-excel_execute", 2)]),
        ])
        self.assertEqual([output["exit_code"] for output in _parse_cli_results(result)], [1, 0, 2])

    def test_rejects_final_namespaced_failure(self):
        result = recorded(execution("excel_execute"), execution("excel-cli-excel_execute", 1))
        with self.assertRaisesRegex(AssertionError, "Final CLI call failed"):
            assert_cli_exit_codes(result)

    def test_strict_mode_rejects_an_earlier_namespaced_failure(self):
        result = recorded(execution("excel-cli-excel_execute", 1), execution("excel_execute"))
        assert_cli_exit_codes(result)
        with self.assertRaisesRegex(AssertionError, "CLI exit codes not zero"):
            assert_cli_exit_codes(result, strict=True)

    def test_default_mode_preserves_failure_rate_check(self):
        result = recorded(*[execution("excel-cli-excel_execute", 1) for _ in range(9)],
                          execution("excel_execute"))
        with self.assertRaisesRegex(AssertionError, "Too many CLI failures: 9/10"):
            assert_cli_exit_codes(result)

    def test_rejects_malformed_namespaced_output(self):
        result = recorded(ToolCall("excel-cli-excel_execute", {"args": "chart read"}, "not JSON"))
        with self.assertRaisesRegex(AssertionError, "Final CLI call failed.*exit_code=-1"):
            assert_cli_exit_codes(result)

    def test_rejects_missing_executions_and_unrelated_tool_names(self):
        for result in (recorded(), recorded(execution("other-server-excel_execute"))):
            with self.subTest(result=result):
                with self.assertRaisesRegex(AssertionError, "No CLI executions recorded"):
                    assert_cli_exit_codes(result)
                with self.assertRaisesRegex(AssertionError, "none did"):
                    assert_cli_args_contain(result, "chart")

    def test_ignores_unrelated_calls_interleaved_with_cli_execution(self):
        result = recorded(execution("other-server-excel_execute", 1), execution("excel-cli-excel_execute"))
        self.assertEqual(len(_parse_cli_results(result)), 1)
        assert_cli_exit_codes(result, strict=True)

    def test_reads_results_from_tool_turns_when_call_results_are_missing(self):
        for name in ("excel_execute", "excel-cli-excel_execute"):
            with self.subTest(name=name):
                result = CopilotResult(turns=[
                    Turn("tool", f"[{name}] {execution(name).result}\n\n" + '{"result":"trace"}'),
                    Turn("assistant", "", [ToolCall(name, {"args": "chart read"})]),
                ])
                self.assertEqual([output["exit_code"] for output in _parse_cli_results(result)], [0])
                assert_cli_exit_codes(result, strict=True)

    def test_does_not_double_count_results_present_in_both_recordings(self):
        call = execution("excel-cli-excel_execute")
        result = CopilotResult(turns=[
            Turn("tool", f"[{call.name}] {call.result}"),
            Turn("assistant", "", [call]),
        ])
        self.assertEqual(len(_parse_cli_results(result)), 1)

    def test_uses_complete_tool_turn_results_when_only_some_calls_have_results(self):
        first = execution("excel_execute")
        second = ToolCall("excel-cli-excel_execute", {"args": "chart read"})
        result = CopilotResult(turns=[
            Turn("tool", f"[{first.name}] {first.result}"),
            Turn("assistant", "", [first]),
            Turn("tool", f"[{second.name}] {execution(second.name, 1).result}"),
            Turn("assistant", "", [second]),
        ])
        with self.assertRaisesRegex(AssertionError, "Final CLI call failed"):
            assert_cli_exit_codes(result)

    def test_rejects_incomplete_tool_turn_recording(self):
        result = recorded(execution("excel_execute"),
                          ToolCall("excel-cli-excel_execute", {"args": "chart read"}))
        with self.assertRaisesRegex(AssertionError, "CLI execution results are incomplete"):
            assert_cli_exit_codes(result)

    def test_does_not_read_unrelated_tool_turns_as_cli_results(self):
        result = CopilotResult(turns=[
            Turn("tool", f"[other-server-excel_execute] {execution('excel_execute').result}"),
            Turn("assistant", "", [ToolCall("excel-cli-excel_execute", {"args": "chart read"})]),
        ])
        with self.assertRaisesRegex(AssertionError, "CLI execution results are incomplete"):
            assert_cli_exit_codes(result)


if __name__ == "__main__":
    unittest.main()
