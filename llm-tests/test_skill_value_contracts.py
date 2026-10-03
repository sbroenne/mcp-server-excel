"""Offline regression checks for evaluation command contracts."""

import hashlib
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from skill_value_tasks import SkillTask


class SkillValueCommandContracts(unittest.TestCase):
    def test_bulk_update_preparation_uses_current_calculation_settings(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "workbook.xlsx"
            path.write_bytes(b"offline fixture")
            task = SkillTask(path, "test", Path("not-an-executable"))
            task.cli = Mock(return_value={"sessionId": "owned"})
            with patch("skill_value_tasks.read_saved_workbook", return_value={"calculationMode": "manual"}):
                task.prepare("bulk-update")
            task.cli.assert_any_call(
                "calculationmode", "set-settings", "--session", "owned", "--mode", "manual",
            )

    def test_read_only_verification_uses_current_calculation_settings_response(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "workbook.xlsx"
            path.write_bytes(b"offline fixture")
            task = SkillTask(path, "test", Path("not-an-executable"))
            task.original_hash = hashlib.sha256(path.read_bytes()).hexdigest()
            task.session = "owned"
            responses = {
                ("session", "list"): {"sessions": [{"sessionId": "owned"}]},
                ("range", "get-values"): {"values": [["Unsaved user note"]]},
                ("calculationmode", "get-settings"): {
                    "mode": "manual", "modeValue": -4135, "settingsScope": "application",
                },
                ("window", "get-info"): {"isVisible": False},
                ("workbook", "get-info"): {"saved": False},
            }
            task.cli = Mock(side_effect=lambda *args: responses[args[:2]])
            result = SimpleNamespace(
                success=True, error=None, evidence_complete=True, capture_errors=[],
                model_used="gpt-6.1-sol",
                final_response="The total formula is incorrect; the correct total is 1630.",
            )
            with patch("skill_value_tasks.assert_read_only") as check_reads:
                task.verify("read-only-audit", result, "cli")
                check_reads.assert_called_once_with(result)
                responses[("calculationmode", "get-settings")].update({"mode": "automatic", "modeValue": -4105})
                with self.assertRaises(AssertionError):
                    task.verify("read-only-audit", result, "cli")
            task.cli.assert_any_call("calculationmode", "get-settings", "--session", "owned")
