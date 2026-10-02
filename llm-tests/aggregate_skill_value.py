"""Export compact, shareable evidence without copying private execution traces."""

from __future__ import annotations

import argparse
import json
from collections import defaultdict
from pathlib import Path
from typing import Any


def _skill_was_read(record: dict[str, Any]) -> bool:
    return any(read.get("success") is True for read in record.get("skill_reads", []))


def aggregate(records: list[dict[str, Any]]) -> dict[str, Any]:
    groups: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    seen: set[tuple[str, str, str, int]] = set()
    cases = []
    for record in records:
        if record.get("passed") is not True:
            continue
        key = (record["task"], record["transport"], record["condition"], record["repetition"])
        if key in seen:
            raise ValueError(f"Duplicate verified case: {key}")
        seen.add(key)
        groups[(record["transport"], record["condition"])].append(record)
        cases.append({
            "task": key[0], "transport": key[1], "condition": key[2], "repetition": key[3],
            "tokens": record.get("tokens"), "skill_read": _skill_was_read(record),
        })
    rows = []
    for (transport, condition), completed in sorted(groups.items()):
        tokens = [record.get("tokens") for record in completed]
        rows.append({
            "transport": transport, "condition": condition, "verified": len(completed),
            "recorded_tokens": sum(tokens) if all(value is not None for value in tokens) else None,
            "skill_reads": sum(_skill_was_read(record) for record in completed),
        })
    effects = {}
    exclusions = {}
    for transport in ("mcp", "cli"):
        baseline = next((row for row in rows if row["transport"] == transport and row["condition"] == "without-skill"), None)
        treatment = next((row for row in rows if row["transport"] == transport and row["condition"] == "with-skill"), None)
        if baseline is None and treatment is None:
            continue
        if baseline is None or treatment is None:
            exclusions[transport] = "Both conditions need verified cases."
            continue
        baseline_cases = {
            (record["task"], record["repetition"])
            for record in groups[(transport, "without-skill")]
        }
        treatment_cases = {
            (record["task"], record["repetition"])
            for record in groups[(transport, "with-skill")]
        }
        if baseline_cases != treatment_cases:
            exclusions[transport] = "Verified task/repetition sets differ between conditions."
            continue
        if baseline["recorded_tokens"] is None or treatment["recorded_tokens"] is None:
            exclusions[transport] = "Recorded usage is missing in at least one condition."
            continue
        if baseline["recorded_tokens"] == 0:
            exclusions[transport] = "Baseline recorded usage is zero; a percentage is undefined."
            continue
        effects[transport] = round(100 * (treatment["recorded_tokens"] / baseline["recorded_tokens"] - 1), 1)
    return {
        "reserved_attempts": len(records), "verified_cases": len(seen),
        "incomplete_or_unverified_attempts": len(records) - len(seen),
        "groups": rows, "token_increase_percent": effects,
        "token_comparison_exclusions": exclusions,
        "cases": sorted(cases, key=lambda row: (row["task"], row["transport"], row["condition"], row["repetition"])),
        "usage_scope": (
            "Verified completed cases only. Interrupted usage is unknown, not zero. "
            "Percentages require matching verified task/repetition sets."
        ),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directories", nargs="+", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    records = []
    manifests = []
    for directory in args.directories:
        manifest = json.loads((directory / "manifest.json").read_text(encoding="utf-8"))
        manifests.append(manifest)
        for path in sorted(directory.glob("[0-9][0-9]-*.json")):
            records.append(json.loads(path.read_text(encoding="utf-8")))
    for manifest in manifests[1:]:
        for field in ("model", "packages", "skills", "suite"):
            if manifest[field] != manifests[0][field]:
                raise ValueError(f"Cannot combine changed {field}")
    receipt = aggregate(records)
    receipt.update({
        "model": manifests[0]["model"], "packages": manifests[0]["packages"],
        "suite": manifests[0]["suite"], "source_commit": manifests[0]["commit"],
        "skill_hashes": manifests[0]["skills"],
        "invocations": [
            {"prior_attempts": manifest["prior_attempts"], "checker_hashes": manifest["checks"]}
            for manifest in manifests
        ],
        "limitations": [
            "One capable model, two repetitions, and no observed correctness difference.",
            "Resumed invocations cross a recorder-code version boundary.",
            "The configured event callback received zero live events; live journaling is unverified.",
            "Recorded tokens are not invoice cost; incomplete attempt usage is unknown.",
            "Public benchmark problems may have appeared in model training.",
        ],
    })
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(receipt, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
