"""No-model regressions for the external dataset loader and outcome checks."""

from __future__ import annotations

import copy
import hashlib
import io
import subprocess
import tarfile
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import spreadsheetbench as bench
from workbook_assertions import read_saved_workbook


def snapshot(values, formulas=None):
    return {
        "sheets": [{
            "name": "Sheet1", "sourceRow": 1, "sourceColumn": 1,
            "sourceValues": values, "sourceFormulas": formulas or copy.deepcopy(values),
            "sourceFormats": [["General"] * len(row) for row in values],
            "tables": [], "charts": [], "pivots": [],
        }],
        "calculationMode": -4105, "queries": [], "slicers": [],
    }


class SpreadsheetBenchChecks(unittest.TestCase):
    def fixture(self):
        before = snapshot([["Keep", 7], [None, None]])
        expected = snapshot([["Keep", 7], [None, 19]])
        spec = bench.CaseSpec("test", "Sheet1", "B2", "B2")
        return spec, before, expected

    def test_correct_answer_passes_but_input_and_partial_answers_fail(self):
        spec, before, expected = self.fixture()
        bench.check_snapshot(spec, expected, before, expected)
        with self.assertRaisesRegex(AssertionError, "Answer"):
            bench.check_snapshot(spec, before, before, expected)
        partial = copy.deepcopy(expected)
        partial["sheets"][0]["sourceValues"][1][1] = None
        with self.assertRaisesRegex(AssertionError, "Answer"):
            bench.check_snapshot(spec, partial, before, expected)

    def test_correct_answer_does_not_excuse_source_corruption_or_extra_output(self):
        spec, before, expected = self.fixture()
        for row, column, value in ((0, 0, "Changed"), (1, 0, "Unexpected output")):
            after = copy.deepcopy(expected)
            after["sheets"][0]["sourceValues"][row][column] = value
            with self.subTest(row=row), self.assertRaisesRegex(AssertionError, "Preservation"):
                bench.check_snapshot(spec, after, before, expected)

    def test_formula_replacement_and_source_format_changes_are_rejected(self):
        spec, before, expected = self.fixture()
        before["sheets"][0]["sourceFormulas"][0][1] = "=3+4"
        expected["sheets"][0]["sourceFormulas"][0][1] = "=3+4"
        for field, value, row, column in (
            ("sourceFormulas", 7, 0, 1), ("sourceFormats", "0.00", 0, 1),
            ("sourceFormats", "yyyy-mm-dd", 1, 0),
        ):
            after = copy.deepcopy(expected)
            after["sheets"][0][field][row][column] = value
            with self.subTest(field=field, row=row), self.assertRaisesRegex(AssertionError, "Preservation"):
                bench.check_snapshot(spec, after, before, expected)

    def test_tolerant_numbers_do_not_confuse_booleans_or_strings(self):
        spec, before, expected = self.fixture()
        for value in ("19", True, 19.01, float("nan")):
            after = copy.deepcopy(expected)
            after["sheets"][0]["sourceValues"][1][1] = value
            with self.subTest(value=value), self.assertRaisesRegex(AssertionError, "Answer"):
                bench.check_snapshot(spec, after, before, expected)
        after["sheets"][0]["sourceValues"][1][1] = 19.000000001
        bench.check_snapshot(spec, after, before, expected)

    def test_existing_sheets_and_objects_cannot_disappear(self):
        spec, before, expected = self.fixture()
        for change in ("sheet", "table", "mode"):
            old, answer, after = (copy.deepcopy(s) for s in (before, expected, expected))
            if change == "sheet":
                after["sheets"][0]["name"] = "Replacement"
            elif change == "table":
                old["sheets"][0]["tables"] = [{"name": "Existing"}]
                answer["sheets"][0]["tables"] = [{"name": "Existing"}]
            else:
                after["calculationMode"] = -4135
            with self.subTest(change=change), self.assertRaises(AssertionError):
                bench.check_snapshot(spec, after, old, answer)

    def test_reference_uses_official_answer_cells_not_incidental_source_format_changes(self):
        spec, before, golden = self.fixture()
        before["sheets"][0]["sourceFormats"][0][1] = "0.00"
        expected = bench.reference_snapshot(spec, before, golden)
        self.assertEqual(expected["sheets"][0]["sourceValues"][1][1], 19)
        self.assertEqual(expected["sheets"][0]["sourceFormats"][0][1], "0.00")
        bench.check_snapshot(spec, expected, before, expected)
        golden["sheets"][0]["sourceValues"][0][1] = 0
        with self.assertRaisesRegex(AssertionError, "Preservation"):
            bench.reference_snapshot(spec, before, golden)

    def test_used_range_offsets_and_missing_trailing_blanks_are_supported(self):
        spec = bench.CaseSpec("test", "Sheet1", "B2", "B2")
        before = snapshot([[None]])
        after = snapshot([[19]])
        after["sheets"][0].update(sourceRow=2, sourceColumn=2)
        bench.check_snapshot(spec, after, before, after)

    def test_answer_range_is_strictly_bounded(self):
        for address in ("A:A", "A0", "XFE1", "A1048577", "B3:A1", "A1,B1", "A1:A10001"):
            with self.subTest(address=address), self.assertRaises(ValueError):
                bench.region(address)

    def test_inspection_limit_is_a_task_failure_but_other_reader_errors_are_not_hidden(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "workbook.xlsx"
            path.write_bytes(b"Reader call is mocked")
            for error, expected in (
                ("InspectionLimitExceeded: too many cells", AssertionError),
                ("Excel is not available", subprocess.CalledProcessError),
            ):
                failure = subprocess.CalledProcessError(1, "pwsh", stderr=error)
                with self.subTest(error=error), patch("workbook_assertions.subprocess.run", side_effect=failure):
                    with self.assertRaises(expected):
                        read_saved_workbook(str(path), use_used_range=True)

    def test_archive_checksum_is_checked_before_contents(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "dataset.tar.gz"
            path.write_bytes(b"not the pinned archive")
            with self.assertRaisesRegex(ValueError, "checksum"):
                bench.load_case("13-1", path)

    def test_symlinks_and_duplicate_exact_members_are_rejected(self):
        for link in (True, False):
            with tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "archive.tar.gz"
                member = f"{bench.ARCHIVE_ROOT}/dataset.json"
                with tarfile.open(path, "w:gz") as archive:
                    info = tarfile.TarInfo(member)
                    if link:
                        info.type = tarfile.SYMTYPE
                        info.linkname = "../../outside.json"
                        archive.addfile(info)
                    else:
                        data = b"[]"
                        info.size = len(data)
                        archive.addfile(info, io.BytesIO(data))
                        archive.addfile(info, io.BytesIO(data))
                with patch.object(bench, "ARCHIVE_SHA256", hashlib.sha256(path.read_bytes()).hexdigest()):
                    with self.subTest(link=link), self.assertRaises(ValueError):
                        bench.load_case("13-1", path)

    def test_metadata_must_match_the_curated_answer_region(self):
        spec = bench.CASES["13-1"]
        for alteration in ({"answer_position": "A1"}, {"instruction": ""}, {"spreadsheet_path": "../outside"}):
            row = {
                "id": "13-1", "instruction": "Do the task", "spreadsheet_path": "spreadsheet/13-1",
                "answer_position": spec.metadata_address, "instruction_type": "Sheet-Level Manipulation",
            }
            row.update(alteration)
            with self.subTest(alteration=alteration), self.assertRaises(ValueError):
                bench.validate_metadata(spec, row)

    def test_formula_probe_uses_changed_dates_units_and_all_search_terms(self):
        values = [[None] * 9 for _ in range(15)]
        for r in range(2, 15):
            values[r][:3] = [50000, "excluded", 1]
        for r, date, text, units in (
            (2, 44562, "velvet,crepe", 5), (4, 44564, "poly", 4),
            (5, 44565, "velvet,crepe", 7), (6, 44566, "velvet,feather", 9),
        ):
            values[r][:3] = [date, text, units]
        values[3][4:6] = [44562, 44566]
        for r, term in enumerate(("velvet", "crepe", "feather", "poly"), start=3):
            values[r][7] = term
        before = snapshot(values)
        changed, expected = bench.formula_probe(before)
        self.assertEqual(before["sheets"][0]["sourceValues"][6][2], 9)
        self.assertEqual([row[8] for row in expected["sheets"][0]["sourceValues"][3:7]], [26, 7, 19, 4])
        self.assertEqual(changed["sheets"][0]["sourceValues"][3][4], 44564)
        self.assertEqual(changed["sheets"][0]["sourceValues"][6][2], 19)

    def test_business_suite_and_external_suite_keep_separate_balanced_cases(self):
        from test_skill_value import cases, REAL_WORLD_NATIVE_TASKS
        self.assertEqual(len(cases(3)), 60)
        self.assertEqual(len(cases(None)), 60)
        selected = cases(None, "real-world")
        self.assertEqual(len(selected), 48)
        self.assertEqual({row[0] for row in selected}, set(bench.TASKS) | set(REAL_WORLD_NATIVE_TASKS))
        self.assertEqual(len(set(selected)), 48)
        self.assertEqual(len(cases(1, "real-world")), 24)
        self.assertEqual(len(cases(3, "real-world")), 72)
        with self.assertRaises(ValueError):
            cases(1, "unknown")

    def test_external_tasks_cannot_fall_through_to_the_business_audit_fixture(self):
        from skill_value_tasks import SkillTask
        with tempfile.TemporaryDirectory() as directory:
            task = SkillTask(Path(directory) / "workbook.xlsx", "test", Path("not-an-executable"))
            with self.assertRaisesRegex(ValueError, "Unknown business task"):
                task.prepare("spreadsheetbench-13-1")
            self.assertFalse(task.path.exists())

    def test_mixed_suite_routes_native_and_public_cases_to_their_correct_fixtures(self):
        from skill_value_tasks import SkillTask, skill_task
        for name in ("control", "query-recovery", "model-refresh", *bench.TASKS):
            request = SimpleNamespace(node=SimpleNamespace(callspec=SimpleNamespace(params={"task": name})))
            with self.subTest(task=name), patch.object(SkillTask, "close_owned_sessions"), patch.object(SkillTask, "cli") as cli:
                generator = skill_task.__wrapped__(request)
                task = next(generator)
                expected = bench.SpreadsheetBenchTask if name in bench.TASKS else SkillTask
                self.assertIs(type(task), expected)
                self.assertFalse(task.path.exists())
                with self.assertRaises(StopIteration):
                    next(generator)
                cli.assert_called_once_with("service", "stop")


if __name__ == "__main__":
    unittest.main()
