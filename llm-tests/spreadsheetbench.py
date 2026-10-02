"""Pinned, opt-in SpreadsheetBench fixtures; workbook access stays in Excel COM."""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import re
import tarfile
import tempfile
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from skill_value import assert_explicit_save
from skill_value_tasks import SkillTask
from workbook_assertions import _one, read_saved_workbook

REVISION = "49b73a94775fb489063f60ca1865e3a650079a79"
ARCHIVE_ROOT = "spreadsheetbench_verified_400"
ARCHIVE_SHA256 = "10ef893dd29cb13ab97143ea787e68cdc9574a13873ab9a54e50b31dc03fc949"
ARCHIVE_SIZE = 14958255
ARCHIVE_URL = (
    f"https://raw.githubusercontent.com/RUCKBReasoning/SpreadsheetBench/{REVISION}"
    f"/data/{ARCHIVE_ROOT}.tar.gz"
)
CACHE = Path(__file__).resolve().parent / "TestResults" / "external-datasets"
ARCHIVE = CACHE / f"{ARCHIVE_ROOT}.tar.gz"
REQUEST_SCOPE = (
    "Complete the attached workbook for its current data. A reusable macro or automatic "
    "refresh for future data is not required. Use formulas where the request asks for formulas. "
)
PRESERVATION_REQUEST = (
    " Keep cells outside the requested result area, their formulas and number formats, "
    "existing worksheets and analysis objects, and the original calculation mode unchanged. "
    "Save and close the same workbook."
)
PROBE_CHANGES = [
    {"sheet": "Sheet1", "cell": "C7", "value": 19},
    {"sheet": "Sheet1", "cell": "E4", "value": 44564},
]


@dataclass(frozen=True)
class CaseSpec:
    id: str
    sheet: str
    address: str
    metadata_address: str
    require_formulas: bool = False


CASES = {
    "13-1": CaseSpec("13-1", "LISTS", "A3:D32", "A3:D32"),
    "267-21": CaseSpec("267-21", "merging", "C2:D11", "merging'!C2:D11"),
    "280-17": CaseSpec("280-17", "Sheet1", "A1:B12", "A1:B12"),
    "38823": CaseSpec("38823", "Sheet1", "I4:I7", "I4:I7", require_formulas=True),
}
TASKS = tuple(f"spreadsheetbench-{id}" for id in CASES)


@dataclass(frozen=True)
class LoadedCase:
    spec: CaseSpec
    instruction: str
    input_bytes: bytes
    answer_bytes: bytes

    def manifest(self) -> dict[str, Any]:
        return {
            "id": self.spec.id, "sheet": self.spec.sheet, "answer_region": self.spec.address,
            "input_sha256": hashlib.sha256(self.input_bytes).hexdigest(),
            "answer_sha256": hashlib.sha256(self.answer_bytes).hexdigest(),
            "instruction_sha256": hashlib.sha256(self.instruction.encode("utf-8")).hexdigest(),
            "require_formulas": self.spec.require_formulas,
            "formula_probe": PROBE_CHANGES if self.spec.id == "38823" else None,
        }


def _checksum(path: Path) -> None:
    if not path.is_file():
        raise FileNotFoundError("Dataset is not cached. Run: uv run python spreadsheetbench.py fetch")
    if hashlib.sha256(path.read_bytes()).hexdigest() != ARCHIVE_SHA256:
        raise ValueError("SpreadsheetBench archive checksum does not match the pinned source")


def _member(archive: tarfile.TarFile, name: str) -> bytes:
    matches = [member for member in archive.getmembers() if member.name == name]
    if len(matches) != 1 or not matches[0].isfile():
        raise ValueError(f"Expected one regular archive member: {name}")
    member = matches[0]
    if not 0 < member.size <= 5_000_000:
        raise ValueError(f"Unexpected archive member size: {name}")
    stream = archive.extractfile(member)
    if stream is None:
        raise ValueError(f"Archive member is not readable: {name}")
    with stream:
        content = stream.read(member.size + 1)
    if len(content) != member.size:
        raise ValueError(f"Incomplete archive member: {name}")
    return content


def validate_metadata(spec: CaseSpec, row: dict[str, Any]) -> None:
    if (
        str(row.get("id")) != spec.id
        or row.get("spreadsheet_path") != f"spreadsheet/{spec.id}"
        or row.get("answer_position") != spec.metadata_address
        or row.get("instruction_type") not in {"Cell-Level Manipulation", "Sheet-Level Manipulation"}
        or not isinstance(row.get("instruction"), str)
        or not row["instruction"].strip()
    ):
        raise ValueError(f"Dataset metadata does not match curated case {spec.id}")


def load_case(id: str, path: Path = ARCHIVE) -> LoadedCase:
    spec = CASES[id]
    _checksum(path)
    with tarfile.open(path, "r:gz") as archive:
        rows = json.loads(_member(archive, f"{ARCHIVE_ROOT}/dataset.json"))
        if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
            raise ValueError("Dataset metadata must be a list of case objects")
        matches = [row for row in rows if str(row.get("id")) == id]
        if len(matches) != 1:
            raise ValueError(f"Expected exactly one dataset record for {id}")
        row = matches[0]
        validate_metadata(spec, row)
        prefix = f"{ARCHIVE_ROOT}/spreadsheet/{id}"
        prompt = _member(archive, f"{prefix}/prompt.txt").decode("utf-8").strip()
        if prompt != row["instruction"].strip():
            raise ValueError(f"Prompt and dataset metadata disagree for {id}")
        return LoadedCase(
            spec, prompt,
            _member(archive, f"{prefix}/1_{id}_init.xlsx"),
            _member(archive, f"{prefix}/1_{id}_golden.xlsx"),
        )


def dataset_manifest() -> dict[str, Any]:
    return {
        "source": "https://github.com/RUCKBReasoning/SpreadsheetBench", "revision": REVISION,
        "archive_sha256": ARCHIVE_SHA256, "license": "CC BY-SA 4.0",
        "scope": "Curated current-workbook pilot, not an official full-benchmark score",
        "request_prefix": REQUEST_SCOPE, "request_suffix": PRESERVATION_REQUEST,
        "cases": [load_case(id).manifest() for id in CASES],
    }


def fetch() -> Path:
    if ARCHIVE.exists():
        _checksum(ARCHIVE)
        return ARCHIVE
    CACHE.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=CACHE, prefix="spreadsheetbench-", suffix=".part", delete=False) as file:
        temporary = Path(file.name)
    try:
        with urllib.request.urlopen(ARCHIVE_URL, timeout=60) as response, temporary.open("wb") as file:
            size = 0
            while chunk := response.read(1024 * 1024):
                size += len(chunk)
                if size > ARCHIVE_SIZE:
                    raise ValueError("Dataset download exceeds the pinned archive size")
                file.write(chunk)
        if size != ARCHIVE_SIZE:
            raise ValueError("Dataset download size does not match the pinned archive")
        _checksum(temporary)
        temporary.replace(ARCHIVE)
    finally:
        temporary.unlink(missing_ok=True)
    return ARCHIVE


def region(address: str) -> set[tuple[int, int]]:
    match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)(?::([A-Z]+)([1-9][0-9]*))?", address)
    if not match:
        raise ValueError(f"Expected a bounded rectangular answer region: {address}")

    def column(text: str) -> int:
        result = 0
        for char in text:
            result = result * 26 + ord(char) - ord("A") + 1
        return result

    left, top = column(match[1]), int(match[2])
    right, bottom = column(match[3] or match[1]), int(match[4] or match[2])
    if not (left <= right <= 16384 and top <= bottom <= 1048576):
        raise ValueError(f"Invalid Excel answer region: {address}")
    if (right - left + 1) * (bottom - top + 1) > 10000:
        raise ValueError(f"Answer region exceeds the inspection limit: {address}")
    return {(r, c) for r in range(top, bottom + 1) for c in range(left, right + 1)}


def _cells(sheet: dict[str, Any], field: str = "sourceValues") -> dict[tuple[int, int], Any]:
    return {
        (sheet["sourceRow"] + r, sheet["sourceColumn"] + c): value
        for r, row in enumerate(sheet[field]) for c, value in enumerate(row)
    }


def _equal(actual: Any, expected: Any) -> bool:
    if expected is None or expected == "":
        return actual is None or actual == ""
    if isinstance(expected, bool) or isinstance(actual, bool):
        return type(actual) is type(expected) and actual == expected
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        return math.isfinite(actual) and math.isfinite(expected) and math.isclose(
            actual, expected, rel_tol=1e-9, abs_tol=1e-8,
        )
    return type(actual) is type(expected) and actual == expected


def _assert_preserved_content(spec: CaseSpec, after: dict[str, Any], before: dict[str, Any]) -> None:
    wanted = region(spec.address)
    for old in before["sheets"]:
        current = _one(after["sheets"], name=old["name"])
        allowed = wanted if old["name"] == spec.sheet else set()
        for field in ("sourceValues", "sourceFormulas"):
            old_cells, current_cells = _cells(old, field), _cells(current, field)
            for coordinate in (old_cells.keys() | current_cells.keys()) - allowed:
                assert _equal(current_cells.get(coordinate), old_cells.get(coordinate)), (
                    f"Preservation: {field} changed on {old['name']} at row/column {coordinate}"
                )


def reference_snapshot(
    spec: CaseSpec, before: dict[str, Any], golden: dict[str, Any],
) -> dict[str, Any]:
    _assert_preserved_content(spec, golden, before)
    reference = copy.deepcopy(before)
    output = _one(reference["sheets"], name=spec.sheet)
    answer = _one(golden["sheets"], name=spec.sheet)
    for field in ("sourceValues", "sourceFormulas", "sourceFormats"):
        answer_cells = _cells(answer, field)
        for row, column in region(spec.address):
            r, c = row - output["sourceRow"], column - output["sourceColumn"]
            assert r >= 0 and c >= 0 and r < len(output[field]) and c < len(output[field][r]), (
                "Reference answer falls outside the curated input's inspected range"
            )
            default = "General" if field == "sourceFormats" else None
            output[field][r][c] = answer_cells.get((row, column), default)
    return reference


def check_snapshot(
    spec: CaseSpec, after: dict[str, Any], before: dict[str, Any], expected: dict[str, Any],
) -> None:
    assert [sheet["name"] for sheet in after["sheets"]] == [sheet["name"] for sheet in before["sheets"]], (
        "Preservation: worksheets or their order changed"
    )
    assert after["calculationMode"] == before["calculationMode"], "Preservation: calculation mode changed"
    wanted = region(spec.address)
    result_sheet = _one(after["sheets"], name=spec.sheet)
    answer_sheet = _one(expected["sheets"], name=spec.sheet)
    actual_values, answer_values = _cells(result_sheet), _cells(answer_sheet)
    actual_formulas = _cells(result_sheet, "sourceFormulas")
    for coordinate in sorted(wanted):
        assert _equal(actual_values.get(coordinate), answer_values.get(coordinate)), (
            f"Answer mismatch on {spec.sheet} at row/column {coordinate}"
        )
        if spec.require_formulas:
            assert str(actual_formulas.get(coordinate, "")).startswith("="), (
                f"Answer must remain a formula on {spec.sheet} at row/column {coordinate}"
            )
    _assert_preserved_content(spec, after, before)
    for old in before["sheets"]:
        current = _one(after["sheets"], name=old["name"])
        allowed = wanted if old["name"] == spec.sheet else set()
        old_formats, current_formats = _cells(old, "sourceFormats"), _cells(current, "sourceFormats")
        for coordinate in (old_formats.keys() | current_formats.keys()) - allowed:
            assert current_formats.get(coordinate, "General") == old_formats.get(coordinate, "General"), (
                f"Preservation: number format changed on {old['name']} at row/column {coordinate}"
            )
        for field in ("tables", "charts", "pivots"):
            for object in old[field]:
                assert object in current[field], f"Preservation: existing {field} changed on {old['name']}"
    for field in ("queries", "slicers"):
        for object in before[field]:
            assert object in after[field], f"Preservation: existing {field} changed"


def formula_probe(before: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    changed = copy.deepcopy(before)
    sheet = _one(changed["sheets"], name="Sheet1")
    for change in PROBE_CHANGES:
        coordinate, = region(change["cell"])
        r, c = coordinate[0] - sheet["sourceRow"], coordinate[1] - sheet["sourceColumn"]
        sheet["sourceValues"][r][c] = change["value"]
        old_formula = sheet["sourceFormulas"][r][c]
        sheet["sourceFormulas"][r][c] = str(change["value"]) if isinstance(old_formula, str) else change["value"]
    values = _cells(sheet)
    start, finish = values[(4, 5)], values[(4, 6)]
    expected = copy.deepcopy(changed)
    output = _one(expected["sheets"], name="Sheet1")
    for row in range(4, 8):
        term = values[(row, 8)]
        total = sum(
            values[(source_row, 3)] for source_row in range(3, 16)
            if start <= values[(source_row, 1)] <= finish
            and term.casefold() in values[(source_row, 2)].casefold()
        )
        output["sourceValues"][row - output["sourceRow"]][9 - output["sourceColumn"]] = total
    return changed, expected


class SpreadsheetBenchTask(SkillTask):
    loaded: LoadedCase | None = None
    expected: dict[str, Any] | None = None

    def prepare(self, task: str) -> None:
        if task not in TASKS:
            raise ValueError(f"Unknown external task: {task}")
        self.loaded = load_case(task.removeprefix("spreadsheetbench-"))
        self.path.write_bytes(self.loaded.input_bytes)
        self.original_hash = hashlib.sha256(self.path.read_bytes()).hexdigest()
        self.before = read_saved_workbook(str(self.path), use_used_range=True, recalculate=True)
        with tempfile.TemporaryDirectory(prefix="excel-bench-answer-") as directory:
            answer = Path(directory) / "answer.xlsx"
            answer.write_bytes(self.loaded.answer_bytes)
            golden = read_saved_workbook(str(answer), use_used_range=True, recalculate=True)
            self.expected = reference_snapshot(self.loaded.spec, self.before, golden)
        check_snapshot(self.loaded.spec, self.expected, self.before, self.expected)

    def prompt(self, task: str) -> str:
        assert self.loaded and task == f"spreadsheetbench-{self.loaded.spec.id}"
        return f"Use the workbook at {self.path}. {REQUEST_SCOPE}{self.loaded.instruction}{PRESERVATION_REQUEST}"

    def verify(self, task: str, result: Any, transport: str) -> dict[str, Any]:
        assert self.loaded and self.before and self.expected
        assert task == f"spreadsheetbench-{self.loaded.spec.id}"
        assert result.success, result.error
        assert result.evidence_complete, result.capture_errors
        assert result.model_used == "gpt-6.1-sol", f"Unexpected model: {result.model_used}"
        assert_explicit_save(result, transport)
        if transport == "cli":
            assert self.cli("session", "list")["sessions"] == [], "Workbook is still open"
        saved_hash = hashlib.sha256(self.path.read_bytes()).hexdigest()
        after = read_saved_workbook(str(self.path), use_used_range=True, recalculate=True)
        check_snapshot(self.loaded.spec, after, self.before, self.expected)
        probed = False
        if self.loaded.spec.id == "38823":
            changed, expected = formula_probe(self.before)
            probe = read_saved_workbook(
                str(self.path), use_used_range=True, recalculate=True, probe_changes=PROBE_CHANGES,
            )
            check_snapshot(self.loaded.spec, probe, changed, expected)
            probed = True
        assert hashlib.sha256(self.path.read_bytes()).hexdigest() == saved_hash, "Inspection changed the saved workbook"
        return {"saved_workbook_verified": True, "source_case": self.loaded.manifest(),
                "formula_probe_passed": probed, "snapshot": after}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Cache the pinned public SpreadsheetBench pilot (no model calls).")
    parser.add_argument("action", choices=("fetch", "manifest"))
    arguments = parser.parse_args()
    if arguments.action == "fetch":
        print(fetch())
    else:
        print(json.dumps(dataset_manifest(), indent=2, ensure_ascii=True))
