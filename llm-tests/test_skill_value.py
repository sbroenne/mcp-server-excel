"""Opt-in, independently scored skill/no-skill comparison within an approved ceiling."""

from __future__ import annotations

import hashlib
import importlib.metadata
import json
import subprocess
from pathlib import Path

import pytest

from conftest import DEFAULT_MAX_TOOL_CALLS, DEFAULT_TIMEOUT_S, build_excel_cli_eval, build_excel_mcp_eval
from consent_scenarios import isolated_cli_servers
from skill_value import check_attempt_budget, paired_results, record_execution, summarize
from skill_discovery import discover_skills, assert_skill_exposure
from skill_value_tasks import FORMATTING_TASKS, TASKS, skill_task

pytestmark = [pytest.mark.aitest, pytest.mark.copilot, pytest.mark.skill_value]
ROOT = Path(__file__).resolve().parents[1]
REAL_WORLD_NATIVE_TASKS = ("query-recovery", "model-refresh")


def cases(repetitions: int | None, suite: str = "business"):
    if suite == "formatting":
        tasks = (*FORMATTING_TASKS, "read-only-audit", "model-refresh")
        repetitions = 2 if repetitions is None else repetitions
    elif suite == "real-world":
        from spreadsheetbench import TASKS as external_tasks
        tasks = (*external_tasks, *REAL_WORLD_NATIVE_TASKS)
        repetitions = 2 if repetitions is None else repetitions
    elif suite == "business":
        tasks = TASKS
        repetitions = 3 if repetitions is None else repetitions
    else:
        raise ValueError(f"Unknown skill-value suite: {suite}")
    return [
        (task, transport, condition, repetition)
        for repetition in range(repetitions)
        for task in tasks
        if suite != "formatting" or task in FORMATTING_TASKS or repetition == 0
        for transport in ("mcp", "cli")
        for condition in (
            ("without-skill", "with-skill") if repetition % 2 == 0 else ("with-skill", "without-skill")
        )
    ]


def pytest_generate_tests(metafunc):
    selected = cases(metafunc.config.getoption("--skill-value-repetitions"),
                     metafunc.config.getoption("--skill-value-suite"))
    metafunc.parametrize(
        "task,transport,condition,repetition", selected,
        ids=[f"{task}-{transport}-{condition}-rep{repetition + 1}"
             for task, transport, condition, repetition in selected],
    )


def hashes(directory: Path) -> dict[str, str]:
    return {str(path.relative_to(directory)): hashlib.sha256(path.read_bytes()).hexdigest()
            for path in sorted(directory.rglob("*")) if path.is_file()}


@pytest.fixture(scope="session")
async def comparison_evidence(request, excel_mcp_skill_dir, excel_cli_skill_dir):
    fields = ("task", "transport", "condition", "repetition")
    suite = request.config.getoption("--skill-value-suite")
    configured = cases(request.config.getoption("--skill-value-repetitions"), suite)
    planned = [tuple(item.callspec.params[field] for field in fields)
               for item in request.session.items
               if item.get_closest_marker("skill_value") and hasattr(item, "callspec")]
    prior = request.config.getoption("--skill-value-prior-attempts")
    ceiling = request.config.getoption("--skill-value-ceiling")
    check_attempt_budget(len(planned), prior, ceiling)
    output = request.config.getoption("--skill-value-output")
    assert output, "--skill-value-output must point to a new comparison evidence directory"
    directory = Path(output).resolve()
    assert not directory.exists(), f"Refusing stale comparison evidence: {directory}"
    directory.mkdir(parents=True)
    manifest = {
        "model": "gpt-6.1-sol", "max_tool_calls": DEFAULT_MAX_TOOL_CALLS, "timeout_s": DEFAULT_TIMEOUT_S,
        "packages": {name: importlib.metadata.version(name)
                     for name in ("pytest-skill-engineering", "github-copilot-sdk", "mcp")},
        "ceiling": ceiling, "prior_attempts": prior, "suite": suite,
        "planned_cases": planned, "configured_cases": configured,
        "commit": subprocess.run(["git", "rev-parse", "HEAD"], cwd=ROOT, check=True, capture_output=True, text=True).stdout.strip(),
        "skills": {"mcp": hashes(Path(excel_mcp_skill_dir)), "cli": hashes(Path(excel_cli_skill_dir))},
        "skill_sources": {
            "entries": hashes(ROOT / "skills"),
            "formatting": hashlib.sha256((ROOT / "docs" / "reference" / "report-formatting.md").read_bytes()).hexdigest(),
        },
        "checks": {name: hashlib.sha256(Path(__file__).with_name(name).read_bytes()).hexdigest()
                   for name in ("test_skill_value.py", "skill_value_tasks.py", "skill_value.py", "workbook_assertions.py",
                                "Inspect-SavedWorkbook.ps1", "conftest.py", "skill_discovery.py",
                                "cli_mcp_server.py", "cli_evidence.py", "consent_scenarios.py",
                                "Close-OwnedWorkbook.ps1", "spreadsheetbench.py", "uv.lock")},
        "usage_policy": "Recorded premium requests if available; otherwise complete input+output tokens. Unknown is not zero.",
        "decision_policy": "Better verified correctness, or >=20% less usage without worse correctness/safety; weak/inconsistent evidence is inconclusive.",
    }
    if suite == "real-world":
        from spreadsheetbench import dataset_manifest
        manifest["external_dataset"] = dataset_manifest()
        manifest["native_tasks"] = REAL_WORLD_NATIVE_TASKS
    manifest["skill_discovery"] = {}
    for transport, builder, supplied in (
        ("mcp", build_excel_mcp_eval, excel_mcp_skill_dir),
        ("cli", build_excel_cli_eval, excel_cli_skill_dir),
    ):
        for condition in ("without-skill", "with-skill"):
            agent = builder("discovery", servers={}, working_directory=str(directory / "probe"),
                            skill_dir=supplied if condition == "with-skill" else None)
            discovered = await discover_skills(agent)
            assert_skill_exposure(discovered, f"excel-{transport}-report-formatting" if condition == "with-skill" else None,
                                  supplied if condition == "with-skill" else None)
            manifest["skill_discovery"][f"{transport}-{condition}"] = discovered
    (directory / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    records = []
    yield directory, manifest, records
    summary = {
        "groups": summarize(records), "pairs": paired_results(records),
        "planned_attempts": len(planned), "recorded_attempts": len(records),
        "complete_selected_cases": len(records) == len(planned),
        "complete_matrix": len(records) == len(configured),
        "limitations": "One model and few repetitions; weak/inconsistent effects are inconclusive.",
    }
    (directory / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")


async def test_skill_value(
    copilot_eval, request, comparison_evidence, skill_task,
    excel_mcp_servers, excel_cli_servers, excel_mcp_skill_dir, excel_cli_skill_dir,
    task, transport, condition, repetition,
):
    directory, manifest, records = comparison_evidence
    check_attempt_budget(len(records) + 1, manifest["prior_attempts"], manifest["ceiling"])
    skill_dir = excel_mcp_skill_dir if transport == "mcp" else excel_cli_skill_dir
    assert hashes(Path(skill_dir)) == manifest["skills"][transport], "Skill changed during comparison"
    skill_task.prepare(task)
    if task == "read-only-audit" and transport == "cli":
        skill_task.session = skill_task.cli("session", "open", str(skill_task.path))["sessionId"]
        skill_task.command("calculationmode", "set-settings", "--mode", "manual")
        skill_task.values("Sheet1", "B7", [["Unsaved user note"]])
    builder = build_excel_mcp_eval if transport == "mcp" else build_excel_cli_eval
    servers = excel_mcp_servers if transport == "mcp" else isolated_cli_servers(
        excel_cli_servers, skill_task.pipe, working_directory=str(skill_task.path.parent),
    )
    agent = builder(
        f"{task}-{transport}-{condition}-rep{repetition + 1}", servers=servers,
        skill_dir=skill_dir if condition == "with-skill" else None,
        model="gpt-6.1-sol", working_directory=str(skill_task.path.parent),
    )
    record = {"case": request.node.nodeid, "task": task, "transport": transport,
              "skill_name": f"excel-{transport}-report-formatting",
              "condition": condition, "repetition": repetition + 1, "passed": False,
              "category": "execution", "verification_error": "Execution did not return evidence",
              "tokens": None, "premium_requests": None}
    if manifest["suite"] == "formatting":
        record["expected_skill_read"] = condition == "with-skill" and task in FORMATTING_TASKS
    records.append(record)
    record_path = directory / f"{len(records):02d}-{task}-{transport}-{condition}-rep{repetition + 1}.json"
    record_path.write_text(json.dumps(record, indent=2), encoding="utf-8")
    request.node.user_properties.append(("skill_value", record))
    try:
        await record_execution(record, record_path, copilot_eval, agent, skill_task.prompt(task),
                               lambda result: skill_task.verify(task, result, transport))
    except Exception as error:
        pytest.exit(f"Comparison stopped: {type(error).__name__}: {error}", returncode=2)
    if record["category"] in {"capture", "execution", "harness"}:
        pytest.exit(f"Comparison stopped: invalid execution/evidence: {record['verification_error']}", returncode=2)
    assert record["passed"], record["verification_error"]
    if manifest["suite"] == "formatting":
        assert record["selection_matches_intent"], (
            "Formatting skill selection did not match the requested task"
        )
