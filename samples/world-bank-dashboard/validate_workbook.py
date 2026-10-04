"""Check real Excel results against the attributed CSVs, using only excelcli."""

import argparse
import csv
import json
import math
from pathlib import Path
import re
import statistics

from build_workbook import ROOT, REGION_COLORS, col
from excel_bridge import Excel


def require(condition, message):
    if not condition:
        raise AssertionError(message)


def near(actual, expected, label):
    require(isinstance(actual, (int, float)) and math.isclose(actual, expected, rel_tol=1e-9, abs_tol=1e-8),
            f"{label}: expected {expected}, received {actual}")


def values(excel, sheet, address):
    result = excel.call("range.get-values", sheetName=sheet, rangeAddress=address)["result"]
    require(not result["cellErrors"], f"Formula errors in {sheet}!{address}: {result['cellErrors']}")
    return result["values"]


def validate_live_licences(excel):
    source = (ROOT / "world_bank_source.m").read_text(encoding="utf-8")
    indicators = re.search(r"IndicatorCodes = (\{.*?\}),", source, re.DOTALL).group(1)
    guard = source.split("LicensesOK = ", 1)[1].split(",\n    Raw = ", 1)[0]
    query = f"""
let
    IndicatorCodes = {indicators},
    GoodRows = List.Transform(IndicatorCodes, each {{_, "CC BY-4.0"}}),
    Check = (rows as list) as logical =>
        let SelectedLicenses = Table.FromRows(rows, {{"Series Code", "License Type"}})
        in {guard},
    Checks = {{
        [Case="Complete", Passed=Check(GoodRows)],
        [Case="Duplicate replaces missing", Passed=not Check(List.RemoveLastN(GoodRows, 1) & {{GoodRows{{0}}}})],
        [Case="Missing", Passed=not Check(List.RemoveLastN(GoodRows, 1))],
        [Case="Restricted", Passed=not Check(List.RemoveLastN(GoodRows, 1) & {{{{List.Last(IndicatorCodes), "Restricted"}}}})]
    }},
    Result = if List.AllTrue(List.Transform(Checks, each [Passed]))
        then Table.FromRecords(Checks)
        else error "Live licence guard accepted missing, duplicate or restricted metadata."
in Result
"""
    excel.call("powerquery.evaluate", mCode=query)


def validate_overview(excel, expected, countries, year=2024):
    def matched(metric):
        return [(expected[(country["Code"], year, metric)], expected[(country["Code"], 2000, metric)])
                for country in countries
                if expected[(country["Code"], year, metric)] is not None
                and expected[(country["Code"], 2000, metric)] is not None]

    income = matched("Income")
    population = matched("Population")
    life = matched("Life")
    near(values(excel, "World Overview", "K11")[0][0],
         statistics.median(a for a, _ in income) / statistics.median(b for _, b in income) - 1,
         "Matched-country income change since 2000")
    near(values(excel, "World Overview", "T11")[0][0],
         sum(a for a, _ in population) / sum(b for _, b in population) - 1,
         "Matched-country population change since 2000")
    near(values(excel, "World Overview", "AC11")[0][0],
         statistics.median(a for a, _ in life) - statistics.median(b for _, b in life),
         "Matched-country life expectancy change since 2000")
    near(values(excel, "World Overview", "AC8")[0][0],
         statistics.median(expected[(country["Code"], year, "Life")] for country in countries
                           if expected[(country["Code"], year, "Life")] is not None),
         "Median life expectancy")
    takeaway = values(excel, "World Overview", "B34")[0][0]
    change = re.search(r"Life expectancy ([+-]?\d+(?:[.,]\d+)?) years", takeaway)
    require(change is not None, f"Missing life expectancy comparison: {takeaway}")
    near(float(change[1].replace(",", ".")),
         round(statistics.median(a for a, _ in life) - statistics.median(b for _, b in life), 1),
         "Displayed life expectancy change")


def validate_overview_chart(excel, expected, countries):
    chart = excel.call("chart.read", chartName="IncomeAndLongevity")["result"]
    require(len(chart["series"]) == len(countries), "Each country needs a named native bubble series")
    for i, (series, country) in enumerate(zip(chart["series"], countries)):
        require(series["name"] == country["Country"], "Chart country names differ from the source")
        require(sum(isinstance(value, (int, float)) for value in series["values"]) == 1,
                f"{country['Country']} must plot exactly one observation")
        near(series["values"][i], expected[(country["Code"], 2024, "Life")], "Bubble vertical position")
        near(series["categories"][i], expected[(country["Code"], 2024, "Income")], "Bubble horizontal position")
        native = excel.call("chartconfig.get-series-settings", chartName="IncomeAndLongevity", seriesIndex=i+1)["result"]
        size_row = i * 2 + 3
        require(f"'Chart Data'!$O${size_row}:$AM${size_row}" in native["formula"],
                "Native bubble-size range must reference its country's population row")
        near(values(excel, "Chart Data", f"{col(15+i)}{size_row}")[0][0],
             expected[(country["Code"], 2024, "Population")] / 1000000, "Bubble population")
        point = excel.call("chartconfig.get-point-format", chartName="IncomeAndLongevity",
                           seriesIndex=i+1, pointIndex=i+1)["result"]
        require(point["fillColor"].upper() == REGION_COLORS[country["Region"]].upper(),
                f"{country['Country']} has the wrong region color")


def validate(excel, logs, live):
    validate_live_licences(excel)
    with (ROOT / "data" / "indicators.csv").open(encoding="utf-8", newline="") as source:
        expected_units = {row["Code"]: row["Unit"] for row in csv.DictReader(source)}
    model_units = excel.call("datamodel.evaluate", daxQuery=(
        'EVALUATE SELECTCOLUMNS(Indicators, "Code", Indicators[Code], "Unit", Indicators[Unit])'
    ))["result"]["rows"]
    require(dict(model_units) == expected_units, "Data Model indicator units differ from the source metadata")
    for sheet in ("Indicators", "Indicator Notes"):
        require({row[0]: row[3] for row in values(excel, sheet, "A2:D10")} == expected_units,
                f"{sheet} units differ from the source metadata")
    with (ROOT / "data" / "observations.csv").open(encoding="utf-8", newline="") as source:
        expected = {(r["CountryCode"], int(r["Year"]), r["Metric"]):
                    float(r["Value"]) if r["Value"] else None for r in csv.DictReader(source)}
    with (ROOT / "data" / "countries.csv").open(encoding="utf-8", newline="") as source:
        countries = list(csv.DictReader(source))
    rows = values(excel, "Model Data", "A2:G5626")
    require(len(rows) == len(expected), "Wrong loaded row count")
    seen = set()
    for code, year, _, metric, value, mode, _ in rows:
        key = (code, int(year), metric)
        require(key not in seen, f"Duplicate loaded key {key}")
        seen.add(key)
        require(mode == "Snapshot", "Shipped workbook must use Snapshot")
        if expected[key] is None:
            require(value is None, f"Missing value was filled: {key}")
        else:
            near(value, expected[key], str(key))
    require(seen == set(expected), "Loaded keys differ from source")

    dax = excel.call("datamodel.evaluate", daxQuery=(
        'EVALUATE SUMMARIZECOLUMNS(Countries[Code], Years[Year], "GDPIndex", [GDP index 2000],'
        ' "Income", [Income], "Population", [Population], "Internet", [Internet])'))["result"]
    require(len(dax["rows"]) == 625, "Expected 25 country x 25 year model results")
    for code, year, index, income, population, internet in dax["rows"]:
        year = int(year)
        near(index, expected[(code, year, "GDP")] / expected[(code, 2000, "GDP")] * 100, f"{code}/{year} index")
        for actual, metric in [(income, "Income"), (population, "Population"), (internet, "Internet")]:
            target = expected[(code, year, metric)]
            if target is None:
                require(actual is None, f"{metric} missing model result was filled")
            else:
                near(actual, target, f"{code}/{year} {metric}")

    validate_overview_chart(excel, expected, countries)
    for name in ("GrowthLines", "InternetLines"):
        info = excel.call("chart.read", chartName=name)["result"]
        require(info["isPivotChart"] and len(info["series"]) == 6, f"{name} must be a six-country native PivotChart")

    income_2024 = [v for (c, y, m), v in expected.items() if y == 2024 and m == "Income" and v is not None]
    near(values(excel, "World Overview", "K8")[0][0], statistics.median(income_2024), "Median country income")
    validate_overview(excel, expected, countries)
    excel.call("slicer.set-slicer-selection", slicerName="OverviewYear", selectedItems=["2000"])
    near(values(excel, "World Overview", "B8")[0][0], 2000, "Year selector")
    validate_overview(excel, expected, countries, 2000)
    excel.call("slicer.set-slicer-selection", slicerName="OverviewYear", selectedItems=["2000", "2024"])
    require(isinstance(values(excel, "World Overview", "T8")[0][0], str), "Multiple years must not report zero population")
    for address in ("K11", "T11", "AC8", "AC11"):
        require(isinstance(values(excel, "World Overview", address)[0][0], str),
                f"Multiple years must not produce a comparison in {address}")
    excel.call("slicer.set-slicer-selection", slicerName="OverviewYear", selectedItems=["2024"])

    excel.call("slicer.set-slicer-selection", slicerName="OverviewRegion", selectedItems=["North America"])
    region = excel.call("chart.read", chartName="IncomeAndLongevity")["result"]
    numeric = [v for series in region["series"] for v in series["values"] if isinstance(v, (float, int)) and v > 0]
    require(len(numeric) == 2, "North America must leave Canada and United States")
    require([row[0] for row in values(excel, "Chart Data", "A2:A26")] ==
            [country["Country"] for country in countries],
            "Country identities must remain fixed when filtering, so region colors stay correct")
    validate_overview(excel, expected, [country for country in countries if country["Region"] == "North America"])
    regions = sorted({r[2] for r in values(excel, "Countries", "A2:D26")})
    excel.call("slicer.set-slicer-selection", slicerName="OverviewRegion", selectedItems=regions)

    selected = ["United States", "China", "Germany", "India", "Brazil", "South Africa"]
    for slicer, sheet in [("GrowthCountries", "Growth & Resilience"), ("ProgressCountries", "Prosperity & Progress")]:
        excel.call("slicer.set-slicer-selection", slicerName=slicer, selectedItems=["Germany"])
        require(values(excel, sheet, "B31")[0][0] == "Germany", f"{sheet} summary did not follow selection")
        require(values(excel, sheet, "B32")[0][0] in ("", None), f"{sheet} retains unselected country")
        excel.call("slicer.set-slicer-selection", slicerName=slicer, selectedItems=selected)

    source = (ROOT / "world_bank_source.m").read_text(encoding="utf-8")
    excel.call("powerquery.update", queryName="WorldBankLive",
               mCode='error "Online source deliberately unavailable in offline check"', refresh=False)
    excel.call("powerquery.refresh", queryName="Observations")
    require(values(excel, "Model Data", "F2")[0][0] == "Snapshot", "Offline refresh failed")
    excel.call("powerquery.update", queryName="WorldBankLive", mCode=source, refresh=False)
    if live:
        excel.call("powerquery.update", queryName="Observations", mCode="WorldBankLive", refresh=True)
        live_rows = values(excel, "Model Data", "A2:G5626")
        require(len(live_rows) == 5625 and all(row[5] == "Live" for row in live_rows), "Live mode did not load")
        require(sum(row[4] is None for row in live_rows) == 46, "Live missing-value coverage changed; review the revised source")
        excel.call("powerquery.update", queryName="Observations", mCode="EmbeddedSnapshot", refresh=True)
    for name in ("SnapshotPivot", "GrowthPivot", "ProgressPivot"):
        excel.call("pivottable.refresh", pivotTableName=name)
    for sheet, stem in [("World Overview", "overview"), ("Growth & Resilience", "growth"),
                        ("Prosperity & Progress", "progress")]:
        values(excel, sheet, "A1:AN39")
        excel.invoke("screenshot", "capture", "--session", excel.session, "--sheet", sheet,
                     "--range", "A1:AN39", "--quality", "High", "--output", logs / f"{stem}.png")
    excel.call("window.set-zoom", sheetName="World Overview", zoom=90)
    return {"loadedObservations": len(rows), "modelCountryYears": len(dax["rows"]),
            "offlineRefresh": "passed", "liveRefresh": "passed" if live else "not run",
            "filters": "passed", "chartBindings": "passed", "formulaErrors": 0}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--cli", required=True)
    parser.add_argument("--workbook", type=Path, default=ROOT / "world-in-motion.xlsx")
    parser.add_argument("--logs", type=Path, required=True)
    parser.add_argument("--live", action="store_true", help="Download the official bulk archive through Excel")
    args = parser.parse_args()
    excel = Excel(args.cli, args.logs)
    excel.open(args.workbook, show=True)
    try:
        report = validate(excel, args.logs.resolve(), args.live)
        (args.logs / "verification.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(report, indent=2))
    finally:
        # Verification exercises filters and queries but never changes the delivered file.
        excel.close(False)


if __name__ == "__main__":
    main()
