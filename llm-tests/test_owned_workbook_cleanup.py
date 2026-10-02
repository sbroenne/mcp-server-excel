"""Unbilled Excel proof that cleanup closes only the fixture and never saves."""

import json
import subprocess
import tempfile
import unittest
from pathlib import Path


class OwnedWorkbookCleanup(unittest.TestCase):
    def test_exact_workbook_cleanup_preserves_peer_and_discards_unsaved_changes(self):
        with tempfile.TemporaryDirectory(prefix="owned-cleanup-") as directory:
            result = subprocess.run(
                ["pwsh", "-NoProfile", "-File", str(Path(__file__).with_name("Test-OwnedWorkbookCleanup.ps1")),
                 "-Directory", directory],
                capture_output=True, text=True, encoding="utf-8", timeout=120,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(json.loads(result.stdout), {
                "closed_only_owned": True, "discarded_changes": True, "repeated_cleanup_safe": True,
            })
