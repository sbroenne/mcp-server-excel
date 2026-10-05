"""Opt-in discovery of installed published CLI skills, without coaching the agent."""

from __future__ import annotations

import copy
import json
import os
import re
import shutil
import subprocess
import sys
import uuid
from dataclasses import replace
from pathlib import Path
from types import SimpleNamespace

import pytest

from conftest import assert_regex, build_excel_cli_eval
from skill_discovery import discover_skills
from workbook_assertions import read_saved_workbook

SKILLS = ("excel-cli", "excel-cli-report-formatting")
VALUES = [
    ["Product", "Quantity", "UnitPrice", "LineTotal"],
    ['Cable "Plus"', 3, 12.5, 37.5],
    ["Monitor", 2, 175.25, 350.5],
    ["Adapter", 4, 8.75, 35],
    ["Total", None, None, 423],
]
FORMULAS = {1: "=B2*C2", 2: "=B3*C3", 3: "=B4*C4", 4: "=SUM(D2:D4)"}
REQUEST = """Create an Excel workbook at {book} with a sheet named Orders.
Put Product, Quantity, UnitPrice, and LineTotal in A1:D1.
The three orders are Cable "Plus" (quantity 3, unit price 12.5),
Monitor (quantity 2, unit price 175.25), and Adapter (quantity 4, unit price 8.75).
Calculate each LineTotal using a formula that multiplies the quantity by the unit price.
Put Total in A5 and a formula adding the three line totals in D5.
Leave B5 and C5 blank. Save and close the workbook, and tell me the grand total."""


def published_agent(plugin: Path, work: Path):
    built = build_excel_cli_eval(
        "published-cli-uncoached-discovery", servers={},
        instructions="Complete the user request using the available tools. Work only in the supplied task folder.",
        working_directory=str(work), allowed_tools=["builtin:*"],
    )
    return replace(
        built, skill_directories=[str(plugin / "skills" / name) for name in SKILLS],
        extra_config={"enable_skills": True},
    )


@pytest.fixture
def published_plugin() -> Path:
    supplied = os.environ.get("EXCEL_PUBLISHED_CLI_PLUGIN")
    if not supplied:
        pytest.skip("Set EXCEL_PUBLISHED_CLI_PLUGIN to an installed published excel-cli plugin.")
    plugin = Path(supplied).resolve()
    assert (plugin / "bin" / "start-cli.ps1").is_file(), "Published plugin launcher is missing"
    metadata = plugin / "plugin.json"
    assert metadata.is_file(), "Published plugin metadata is missing"
    assert json.loads(metadata.read_text(encoding="utf-8"))["name"] == "excel-cli", "Wrong published plugin"
    directories = plugin / "skills"
    assert directories.is_dir(), "Published plugin skills directory is missing"
    assert {path.name for path in directories.iterdir() if path.is_dir()} == set(SKILLS), (
        "Expected exactly the launcher-discovery and report-formatting skills"
    )
    assert all((directories / name / "SKILL.md").is_file() for name in SKILLS), "Published skill source is missing"
    return plugin


async def assert_published_discovery(agent, plugin: Path) -> None:
    discovered = await discover_skills(agent, required_tools=("powershell", "skill"))
    assert len(discovered) == len(SKILLS) and {skill["name"] for skill in discovered} == set(SKILLS), (
        f"Wrong published skill availability: {discovered}"
    )
    for skill in discovered:
        assert skill["enabled"] is True, f"Disabled skill: {skill['name']}"
        assert Path(skill["path"]).resolve() == (plugin / "skills" / skill["name"] / "SKILL.md").resolve(), (
            f"Wrong published skill source: {skill}"
        )


def assert_orders(snapshot) -> None:
    matching = [sheet for sheet in snapshot["sheets"] if sheet["name"] == "Orders"]
    assert len(matching) == 1, "Expected exactly one Orders worksheet"
    sheet = matching[0]
    assert sheet["sourceValues"] == VALUES, "Saved values differ from the request"
    for row, expected in FORMULAS.items():
        assert sheet["sourceFormulas"][row][3] == expected, "Missing or wrong formula"


def assert_launcher_operations(result, launcher: Path, book: Path) -> None:
    calls = result.all_tool_calls
    assert calls and all(call.evidence_complete and call.success is True for call in calls), (
        "Every captured call must have successful correlated completion evidence"
    )
    assert any(call.name == "skill" and "excel-cli" in call.arguments.values() for call in calls), (
        "The agent did not activate the published launcher-discovery skill"
    )
    launcher_calls = [
        call for call in calls if call.name == "powershell"
        and str(launcher).casefold() in str(call.arguments.get("command", "")).casefold()
    ]
    mutations = set()
    closed = False
    for call in launcher_calls:
        command = call.arguments["command"]
        outputs = [
            json.loads(line) for line in (call.result or "").splitlines()
            if line.lstrip().startswith("{")
        ]
        for output in outputs:
            if "success" in output:
                assert output["success"] is True and not output.get("errorMessage"), (
                    f"Packaged launcher operation failed: {output}"
                )
            if output.get("filePath") and Path(output["filePath"]).resolve() == book.resolve():
                if output.get("action") in ("set-values", "set-formulas") and output.get("success") is True:
                    mutations.add(output["action"])
            if (re.search(r"\bsession\s+close\b", command, re.IGNORECASE) and "--save" in command
                    and output.get("success") is True and output.get("message") == "Session closed and saved."):
                closed = True
    assert mutations == {"set-values", "set-formulas"}, "No successful packaged launcher workbook mutations"
    assert closed, "No successful packaged launcher save-and-close operation"


def owned_workbook(book: Path, *, inspect_only: bool = False):
    return subprocess.run(
        ["pwsh", "-NoProfile", "-File", str(Path(__file__).with_name("Close-OwnedWorkbook.ps1")),
         "-Path", str(book), *(["-InspectOnly"] if inspect_only else [])],
        capture_output=True, text=True, encoding="utf-8", timeout=60, check=True,
    )


def owned_command(plugin: Path, pipe: str, *arguments: str):
    return subprocess.run(
        ["pwsh", "-NoProfile", "-File", str(plugin / "bin" / "start-cli.ps1"), "-q", *arguments],
        env={**os.environ, "EXCELMCP_CLI_PIPE": pipe}, capture_output=True,
        text=True, encoding="utf-8", timeout=180, check=True,
    )


@pytest.mark.offline
@pytest.mark.parametrize("mutation", ("missing-sheet", "wrong-sheet", "wrong-values", "constants", "wrong-formula"))
def test_orders_checker_rejects_wrong_state(mutation):
    formulas = copy.deepcopy(VALUES)
    for row, value in FORMULAS.items():
        formulas[row][3] = value
    snapshot = {"sheets": [{"name": "Orders", "sourceValues": copy.deepcopy(VALUES), "sourceFormulas": formulas}]}
    assert_orders(snapshot)
    sheet = snapshot["sheets"][0]
    if mutation == "missing-sheet":
        snapshot["sheets"] = []
    elif mutation == "wrong-sheet":
        sheet["name"] = "Other"
    elif mutation == "wrong-values":
        sheet["sourceValues"][1][1] = 99
    elif mutation == "constants":
        sheet["sourceFormulas"] = copy.deepcopy(VALUES)
    else:
        sheet["sourceFormulas"][4][3] = "=423"
    with pytest.raises(AssertionError):
        assert_orders(snapshot)


@pytest.mark.offline
@pytest.mark.parametrize("mutation", ("none", "help-only", "wrong-wrapper", "failed", "incomplete", "no-skill", "no-close"))
def test_launcher_checker_rejects_unproven_operations(tmp_path, mutation):
    launcher, book = tmp_path / "start-cli.ps1", tmp_path / "orders.xlsx"
    calls = [
        SimpleNamespace(name="skill", arguments={"skill": "excel-cli"}, result="loaded",
                        success=True, evidence_complete=True),
        SimpleNamespace(name="powershell", arguments={"command": f"& '{launcher}' range set-values"},
                        result=json.dumps({"success": True, "filePath": str(book), "action": "set-values"}),
                        success=True, evidence_complete=True),
        SimpleNamespace(name="powershell", arguments={"command": f"& '{launcher}' range set-formulas"},
                        result=json.dumps({"success": True, "filePath": str(book), "action": "set-formulas"}),
                        success=True, evidence_complete=True),
        SimpleNamespace(name="powershell", arguments={"command": f"& '{launcher}' session close --save"},
                        result=json.dumps({"success": True, "message": "Session closed and saved."}),
                        success=True, evidence_complete=True),
    ]
    result = SimpleNamespace(all_tool_calls=calls)
    assert_launcher_operations(result, launcher, book)
    if mutation == "none":
        calls.clear()
    elif mutation == "help-only":
        calls[1].result = "help"
    elif mutation == "wrong-wrapper":
        calls[1].arguments["command"] = "excelcli range set-values"
    elif mutation == "failed":
        calls[1].result = json.dumps({"success": False, "action": "set-values", "filePath": str(book)})
    elif mutation == "incomplete":
        calls[1].evidence_complete = False
    elif mutation == "no-skill":
        calls[0].arguments["skill"] = "excel-cli-report-formatting"
    else:
        calls.pop()
    with pytest.raises(AssertionError):
        assert_launcher_operations(result, launcher, book)


@pytest.mark.offline
def test_published_agent_has_no_cli_bridge_or_custom_workspace(tmp_path):
    agent = published_agent(tmp_path / "plugin", tmp_path / "task")
    assert agent.mcp_servers == {} and agent.allowed_tools == ["builtin:*"]
    assert agent.extra_config == {"enable_skills": True}
    assert agent.skill_directories == [str(tmp_path / "plugin" / "skills" / name) for name in SKILLS]
    assert not any(word in REQUEST.lower() for word in ("cli", "npx", "skill", "start-cli", "powershell"))


@pytest.mark.offline
@pytest.mark.parametrize("mutation", ("missing-launcher", "wrong-plugin", "missing-skill"))
def test_supplied_invalid_plugin_fails(tmp_path, monkeypatch, mutation):
    (tmp_path / "bin").mkdir()
    (tmp_path / "bin" / "start-cli.ps1").touch()
    (tmp_path / "plugin.json").write_text('{"name":"excel-cli"}', encoding="utf-8")
    for name in SKILLS:
        directory = tmp_path / "skills" / name
        directory.mkdir(parents=True)
        (directory / "SKILL.md").touch()
    monkeypatch.setenv("EXCEL_PUBLISHED_CLI_PLUGIN", str(tmp_path))
    assert published_plugin.__wrapped__() == tmp_path.resolve()
    if mutation == "missing-launcher":
        (tmp_path / "bin" / "start-cli.ps1").unlink()
    elif mutation == "wrong-plugin":
        (tmp_path / "plugin.json").write_text('{"name":"excel-mcp"}', encoding="utf-8")
    else:
        (tmp_path / "skills" / SKILLS[0] / "SKILL.md").unlink()
    with pytest.raises(AssertionError):
        published_plugin.__wrapped__()


async def test_published_cli_unpaid_preflight(github_auth, published_plugin, tmp_path):
    await assert_published_discovery(published_agent(published_plugin, tmp_path / "task"), published_plugin)


@pytest.mark.skipif(
    os.environ.get("EXCEL_RUN_PUBLISHED_CLI_DISCOVERY") != "1",
    reason="One paid attempt requires EXCEL_RUN_PUBLISHED_CLI_DISCOVERY=1 and explicit test selection.",
)
@pytest.mark.aitest
@pytest.mark.copilot
async def test_live_published_cli_discovery(copilot_eval, published_plugin, tmp_path, monkeypatch):
    if sys.platform != "win32" or any(shutil.which(command) is None for command in ("pwsh", "node", "npx")):
        pytest.skip("Published CLI discovery requires Windows, PowerShell 7, and Node.js with npm/npx.")
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
            pass
    except FileNotFoundError:
        pytest.skip("Published CLI discovery requires installed desktop Excel.")
    work = tmp_path / "task"
    agent = published_agent(published_plugin, work)
    await assert_published_discovery(agent, published_plugin)
    book = work / "orders.xlsx"
    pipe = f"llm-published-cli-{uuid.uuid4().hex}"
    monkeypatch.setenv("EXCELMCP_CLI_PIPE", pipe)
    monkeypatch.setenv("COPILOT_AUTO_UPDATE", "false")
    try:
        result = await copilot_eval(agent, REQUEST.format(book=book))
        assert result.success, (result.stop_reason, result.error)
        assert result.evidence_complete, result.capture_errors
        assert_launcher_operations(result, published_plugin / "bin" / "start-cli.ps1", book)
        assert json.loads(owned_workbook(book, inspect_only=True).stdout)["matched"] == 0, (
            "The agent left the workbook open before independent verification"
        )
        snapshot = read_saved_workbook(str(book), "A1:D5", recalculate=True)
        (tmp_path / "orders-snapshot.json").write_text(json.dumps(snapshot, indent=2), encoding="utf-8")
        assert_orders(snapshot)
        assert_regex(result.final_response, r"\b423(?:\.00)?\b")
    finally:
        try:
            owned_workbook(book)
        finally:
            if json.loads(owned_command(published_plugin, pipe, "service", "status").stdout)["running"]:
                owned_command(published_plugin, pipe, "service", "stop")
