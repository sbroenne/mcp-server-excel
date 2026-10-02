"""Independent checks of saved workbooks, not of an agent's claimed results."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path
from typing import Any


def read_saved_workbook(
    path: str, source_range: str = "A1:E9", *, include_analysis: bool = False,
    use_used_range: bool = False, recalculate: bool = False,
    probe_changes: list[dict[str, Any]] | None = None,
    include_presentation: bool = False,
) -> dict[str, Any]:
    workbook = Path(path)
    assert workbook.is_file(), f"Workbook was not saved: {workbook}"
    command = [
        "pwsh", "-NoProfile", "-File",
        str(Path(__file__).with_name("Inspect-SavedWorkbook.ps1")),
        "-Path", str(workbook.resolve()), "-SourceRange", source_range,
        *(["-IncludeAnalysis"] if include_analysis else []),
        *(["-IncludePresentation"] if include_presentation else []),
        *(["-UseUsedRange"] if use_used_range else []),
        *(["-Recalculate"] if recalculate else []),
        *(["-ProbeChangesJson", json.dumps(probe_changes)] if probe_changes is not None else []),
    ]
    try:
        inspected = subprocess.run(
            command, capture_output=True, text=True, encoding="utf-8", timeout=180, check=True,
        )
    except subprocess.CalledProcessError as error:
        if "InspectionLimitExceeded:" in (error.stderr or ""):
            raise AssertionError("Saved workbook exceeds the 10000-cell inspection limit") from error
        raise
    return json.loads(inspected.stdout)


def _one(items: list[dict[str, Any]], **expected: Any) -> dict[str, Any]:
    matches = [item for item in items if all(item[key] == value for key, value in expected.items())]
    assert len(matches) == 1, f"Expected exactly one {expected}; found {len(matches)}"
    return matches[0]


def _assert_slicer_position(snapshot: dict[str, Any], slicer: dict[str, Any], sheet: str, address: str) -> None:
    assert slicer["sheet"] == sheet
    target = _one(snapshot["sheets"], name=sheet)["positions"][address]
    # Excel's floating-point shape coordinates can put a boundary in the adjacent anchor cell.
    for axis in ("left", "top"):
        assert abs(slicer[axis] - target[axis]) < 0.1, (slicer, target)


def assert_chart(snapshot: dict[str, Any], *, below: bool) -> None:
    sheet = _one(snapshot["sheets"], name="Sheet1")
    chart = _one(sheet["charts"])
    bounds = sheet["bounds"]
    assert chart["width"] > 0 and chart["height"] > 0
    if below:
        assert chart["top"] >= bounds["top"] + bounds["height"], "Chart overlaps the source rows"
        assert chart["type"] == 51, "Expected a clustered column chart"
        expected = {
            "Revenue": [50000, 55000, 48000, 62000, 58000],
            "Expenses": [35000, 38000, 32000, 41000, 39000],
        }
        categories = ["January", "February", "March", "April", "May"]
    else:
        assert chart["left"] >= bounds["left"] + bounds["width"], "Chart overlaps the source columns"
        assert chart["type"] in (4, 65), "Expected a line chart"
        _one(sheet["tables"], name="ProductSales")
        expected = {"Q1": [100, 80, 200, 50], "Q2": [150, 90, 180, 60], "Q3": [120, 110, 220, 75]}
        categories = ["Widget", "Gadget", "Device", "Tool"]
    assert len(chart["series"]) == len(expected)
    for name, values in expected.items():
        series = _one(chart["series"], name=name)
        assert series["values"] == values, f"Wrong values for {name}"
        assert series["categories"] == categories, f"Wrong categories for {name}"
    expected_cells = [
        ["Month" if below else "Product", *expected],
        *[[category, *[values[index] for values in expected.values()]]
          for index, category in enumerate(categories)],
    ]
    assert sheet["sourceValues"] == expected_cells, "Saved source data differs from the request"


def assert_pivot_slicer(snapshot: dict[str, Any]) -> None:
    slicer = _one(snapshot["slicers"], field="Region")
    assert len(snapshot["slicers"]) == 1, "Temporary Product slicer was not removed"
    assert slicer["selected"] == ["North"]
    _assert_slicer_position(snapshot, slicer, "Analysis", "E2")
    table = _one(_one(snapshot["sheets"], name="Sheet1")["tables"], name="SalesData")
    assert len(table["rows"]) == 8 and sum(row[3] for row in table["rows"]) == 92200
    pivot = _one(_one(snapshot["sheets"], name="Analysis")["pivots"])
    assert slicer["pivots"] == [pivot["name"]]
    assert pivot["values"][-1][-1] == 50500
    labels = [row[0] for row in pivot["values"]]
    assert "North" in labels and "South" not in labels


def assert_table_slicers(snapshot: dict[str, Any]) -> None:
    assert len(snapshot["slicers"]) == 2
    for field, selected, position in (("Department", ["Engineering"], "F2"), ("Status", ["Active"], "H2")):
        slicer = _one(snapshot["slicers"], field=field)
        assert slicer["selected"] == selected
        assert slicer["table"] == "Employees"
        _assert_slicer_position(snapshot, slicer, "Sheet1", position)
    table = _one(_one(snapshot["sheets"], name="Sheet1")["tables"], name="Employees")
    assert len(table["rows"]) == 8
    assert table["visibleRows"] == [
        ["Engineering", "Alice", "Active", 85000],
        ["Engineering", "Bob", "Active", 92000],
        ["Engineering", "Grace", "Active", 88000],
    ]


def assert_combined_slicers(snapshot: dict[str, Any]) -> None:
    assert len(snapshot["slicers"]) == 2
    warehouse = _one(snapshot["slicers"], field="Warehouse")
    category = _one(snapshot["slicers"], field="Category")
    assert warehouse["selected"] == ["West"] and warehouse["table"] == "Inventory"
    assert category["selected"] == ["Electronics"]
    _assert_slicer_position(snapshot, warehouse, "Sheet1", "G2")
    _assert_slicer_position(snapshot, category, "Summary", "D2")
    table = _one(_one(snapshot["sheets"], name="Sheet1")["tables"], name="Inventory")
    assert len(table["rows"]) == 8
    assert table["visibleRows"] == [
        ["Electronics", "Laptop", "West", 50, 999],
        ["Electronics", "Phone", "West", 120, 599],
        ["Furniture", "Desk", "West", 25, 350],
        ["Furniture", "Chair", "West", 40, 175],
    ]
    pivot = _one(_one(snapshot["sheets"], name="Summary")["pivots"])
    assert category["pivots"] == [pivot["name"]]
    assert pivot["values"][-1][-1] == 285, "A Table slicer must not masquerade as a PivotTable filter"
