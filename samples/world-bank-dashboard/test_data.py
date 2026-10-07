"""Small, offline checks for the distributable source snapshot."""

import csv
import json
import math
from pathlib import Path
import re
import tempfile
import unittest
from unittest.mock import patch

from prepare_data import COUNTRIES, INDICATORS
from build_workbook import main, redesign_overview, REGION_COLORS

ROOT = Path(__file__).parent


def read_csv(name):
    with (ROOT / "data" / name).open(encoding="utf-8", newline="") as source:
        return list(csv.DictReader(source))


class SnapshotTests(unittest.TestCase):
    def test_complete_unique_observation_keys(self):
        observations = read_csv("observations.csv")
        keys = [(row["CountryCode"], int(row["Year"]), row["IndicatorCode"]) for row in observations]
        expected = {(country, year, indicator) for country in COUNTRIES
                    for year in range(2000, 2025) for indicator in INDICATORS}
        self.assertEqual(len(keys), len(set(keys)))
        self.assertEqual(set(keys), expected)

    def test_missing_and_finite_values(self):
        observations = read_csv("observations.csv")
        manifest = json.loads((ROOT / "data" / "sources.json").read_text(encoding="utf-8"))
        self.assertEqual(len(observations), manifest["observations"])
        self.assertEqual(sum(not row["Value"] for row in observations), manifest["missing"])
        for row in observations:
            if row["Value"]:
                self.assertTrue(math.isfinite(float(row["Value"])))
            self.assertEqual(row["Metric"], INDICATORS[row["IndicatorCode"]][0])

    def test_indicator_licences_and_definitions(self):
        metadata = read_csv("indicators.csv")
        self.assertEqual({row["Code"] for row in metadata}, set(INDICATORS))
        for row in metadata:
            self.assertEqual(row["License"], "CC BY-4.0")
            self.assertTrue(row["Source"] and row["Definition"])

    def test_country_metadata(self):
        countries = read_csv("countries.csv")
        self.assertEqual(len(countries), len(COUNTRIES))
        self.assertEqual({row["Code"] for row in countries}, set(COUNTRIES))
        self.assertTrue(all(row["Region"] and row["Country"] for row in countries))

    def test_growth_and_inflation_units_are_annual_percent_changes(self):
        metadata = {row["Code"]: row for row in read_csv("indicators.csv")}
        for code in ("NY.GDP.MKTP.KD.ZG", "FP.CPI.TOTL.ZG"):
            self.assertEqual(INDICATORS[code][1], "Annual percent change")
            self.assertEqual(metadata[code]["Unit"], "Annual percent change")

    def test_public_query_matches_extraction_scope(self):
        code = (ROOT / "world_bank_source.m").read_text(encoding="utf-8")
        countries = re.search(r'CountryCodes = Text.Split\("([^"]+)"', code).group(1).split()
        indicators = re.search(r"IndicatorCodes = \{([^}]+)\}", code).group(1)
        self.assertEqual(countries, COUNTRIES)
        self.assertEqual(re.findall(r'"([^"]+)"', indicators), list(INDICATORS))
        self.assertNotIn("Excel.CurrentWorkbook", code)
        self.assertNotIn("File.Contents", code)
        self.assertIn("PreserveMissing", code)

    def test_snapshot_has_no_network_dependency(self):
        code = (ROOT / "snapshot_source.m").read_text(encoding="utf-8")
        self.assertIn("Excel.CurrentWorkbook", code)
        self.assertNotIn("Web.Contents", code)
        self.assertNotIn("File.Contents", code)


class OverviewDesignTests(unittest.TestCase):
    def setUp(self):
        class Recorder:
            def __init__(self):
                self.commands = []

            def batch(self, commands):
                self.commands.extend(commands)

            def call(self, command, **args):
                self.commands.append({"command": command, "args": args})
                if command == "table.list":
                    return {"result": {"tables": [{"name": "OverviewData"}]}}
                return None

        recorder = Recorder()
        redesign_overview(recorder)
        self.commands = recorder.commands

    def test_only_overview_and_its_supporting_data_are_changed(self):
        sheets = {c["args"]["sheetName"] for c in self.commands if "sheetName" in c["args"]}
        self.assertEqual(sheets, {"World Overview", "Chart Data"})
        self.assertFalse(any(c["command"].startswith(("powerquery.", "datamodel.", "datamodelrelationship."))
                             for c in self.commands))

    def test_named_series_keep_population_and_missing_values(self):
        matrix = next(c["args"]["formulas"] for c in self.commands
                      if c["command"] == "range.set-formulas" and c["args"]["rangeAddress"] == "O1:AM51")
        self.assertEqual(len(matrix), 51)
        for i in range(25):
            self.assertEqual(matrix[0][i], f"=B{i+2}")
            self.assertEqual(matrix[2*i+1][i], f"=C{i+2}")
            self.assertEqual(matrix[2*i+2][i], f"=D{i+2}")
            self.assertEqual(sum(value == "=NA()" for value in matrix[2*i+1]), 24)
            self.assertEqual(sum(value == "=NA()" for value in matrix[2*i+2]), 24)

    def test_country_series_colors_cover_every_region(self):
        colors = [c["args"]["fillColor"] for c in self.commands
                  if c["command"] == "chartconfig.set-series-format"]
        countries = read_csv("countries.csv")
        self.assertEqual(colors, [REGION_COLORS[c["Region"]] for c in countries])


class BuildLifecycleTests(unittest.TestCase):
    def test_failed_reopen_keeps_original_error_and_does_not_close_again(self):
        with tempfile.TemporaryDirectory() as directory:
            args = ["build_workbook.py", "--cli", "excelcli.exe",
                    "--output", str(Path(directory) / "sample.xlsx"),
                    "--logs", str(Path(directory) / "logs")]
            responses = [
                {"sessionId": "created"},
                {"sessions": [{"sessionId": "created", "canClose": True}]},
                {"success": True},
                RuntimeError("Cannot reopen saved model"),
            ]
            with patch("sys.argv", args), patch("build_workbook.build_model"), \
                    patch("build_workbook.Excel.invoke", side_effect=responses) as invoke:
                with self.assertRaisesRegex(RuntimeError, "Cannot reopen saved model"):
                    main()
                self.assertEqual(invoke.call_count, 4)


if __name__ == "__main__":
    unittest.main()
