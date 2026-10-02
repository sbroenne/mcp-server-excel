"""Fixtures and helpers for ExcelMcp LLM integration tests."""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path
from typing import Any

import pytest

from pytest_skill_engineering.copilot import CopilotEval
from pytest_skill_engineering.copilot.result import ToolCall
from copilot.tools import Tool, ToolInvocation, ToolResult
from cli_evidence import tool_output

TESTS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TESTS_DIR.parent
FIXTURES_DIR = TESTS_DIR / "Fixtures"
TEST_RESULTS_DIR = TESTS_DIR / "TestResults"
TEST_RESULTS_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_MODEL = os.environ.get("EXCEL_LLM_MODEL", "gpt-6.1-sol")
DEFAULT_MAX_TURNS = 20
DEFAULT_MAX_TOOL_CALLS = 80
DEFAULT_TIMEOUT_S = 600.0

_MCP_INSTRUCTIONS = (
    "You are an Excel automation assistant. Use the available MCP tools to complete the "
    "workbook task end-to-end. Save workbooks when the task asks for persistence."
)
_CLI_INSTRUCTIONS = (
    "You are an Excel CLI automation assistant. Use the excel CLI tool to complete the "
    "workbook task end-to-end. Save workbooks when the task asks for persistence."
)


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption("--run-skill-value", action="store_true", help="Enable paid skill/no-skill experiments.")
    parser.addoption("--skill-value-suite", choices=("formatting", "business", "real-world"), default="formatting",
                     help="Select formatting and non-trigger tasks, business tasks, or public workbook cases.")
    parser.addoption("--skill-value-output", help="New directory for comparison manifest and verification records.")
    parser.addoption("--skill-value-repetitions", type=int, choices=(1, 2, 3), default=None,
                     help="Balanced repetitions: defaults to 3 for business, 2 for formatting/real-world.")
    parser.addoption("--skill-value-prior-attempts", type=int, default=0,
                     help="Attempts already counted against the explicitly approved execution ceiling.")
    parser.addoption("--skill-value-ceiling", type=int, choices=(60, 100), default=60,
                     help="Approved ceiling; select 100 only with explicit authorization.")


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    for item in items:
        fixturenames = set(getattr(item, "fixturenames", []))
        if "copilot_eval" in fixturenames and not any(m.name == "copilot" for m in item.iter_markers()):
            item.add_marker(pytest.mark.copilot)
        if item.get_closest_marker("skill_value") and not config.getoption("--run-skill-value"):
            item.add_marker(pytest.mark.skip(reason="Paid comparison requires --run-skill-value"))


def _has_github_auth() -> bool:
    if os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN"):
        return True
    if shutil.which("gh") is None:
        return False

    try:
        result = subprocess.run(
            ["gh", "auth", "status"],
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False

    return result.returncode == 0


@pytest.fixture(scope="session")
def github_auth() -> None:
    if not _has_github_auth():
        pytest.skip(
            "GitHub auth required for pytest-skill-engineering Copilot tests. "
            "Set GITHUB_TOKEN/GH_TOKEN or run `gh auth login`."
        )


@pytest.fixture(autouse=True)
def live_prerequisites(request: pytest.FixtureRequest) -> None:
    if "copilot_eval" in request.fixturenames:
        request.getfixturevalue("github_auth")


def unique_path(prefix: str, suffix: str = ".xlsx") -> str:
    temp_dir = Path(os.environ.get("TEMP", tempfile.gettempdir()))
    return str(temp_dir / f"{prefix}-{uuid.uuid4()}{suffix}")


def unique_results_path(prefix: str, suffix: str = ".xlsx") -> str:
    return str(TEST_RESULTS_DIR / f"{prefix}-{uuid.uuid4()}{suffix}")


def assert_regex(text: str | None, pattern: str) -> None:
    haystack = text or ""
    if not re.search(pattern, haystack, re.IGNORECASE | re.MULTILINE):
        raise AssertionError(f"Pattern not found: {pattern}\nText:\n{haystack}")


def _cli_tool_calls(result: Any) -> list[ToolCall]:
    return [
        call for call in result.all_tool_calls
        if call.name in {"excel_execute", "excel-cli-excel_execute"}
    ]


def _parse_cli_results(result: Any) -> list[dict[str, Any]]:
    calls = _cli_tool_calls(result)
    if not calls:
        return []

    payloads = [call.result or "" for call in calls]
    use_tool_turns = any(not payload for payload in payloads)
    if use_tool_turns:
        # Some SDK recordings keep outputs only in tool turns, not on ToolCall.
        payloads = []
        for turn in result.turns:
            if turn.role != "tool":
                continue
            content = turn.content or ""
            marker = re.match(r"^\[(excel_execute|excel-cli-excel_execute)\]\s*", content)
            if marker:
                payloads.append(content[marker.end():])
        if len(payloads) != len(calls):
            raise AssertionError(
                f"CLI execution results are incomplete: {len(payloads)}/{len(calls)} recorded"
            )

    outputs: list[dict[str, Any]] = []
    for index, payload in enumerate(payloads):
        try:
            # Tool turns can append SDK trace JSON after the execution object.
            output = json.JSONDecoder().raw_decode(payload)[0] if use_tool_turns else tool_output(calls[index])
            outputs.append(output)
        except json.JSONDecodeError:
            outputs.append({"exit_code": -1, "stdout": payload, "stderr": ""})

    return outputs


def assert_cli_exit_codes(result: Any, *, strict: bool = False) -> None:
    outputs = _parse_cli_results(result)
    if not outputs:
        raise AssertionError("No CLI executions recorded")

    calls = [call for call in result.all_tool_calls if call.name in {"excel_execute", "excel-cli-excel_execute"}]
    if any(not call.evidence_complete for call in calls):
        raise AssertionError("Incomplete CLI execution evidence")
    for output in outputs if strict else outputs[-1:]:
        stdout = output.get("stdout", "")
        for line in stdout.splitlines():
            try:
                data = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(data, dict) and data.get("success") is False:
                raise AssertionError(f"CLI operation failed: {data}")

    if strict:
        failures = [output for output in outputs if output.get("exit_code") != 0]
        if failures:
            raise AssertionError(f"CLI exit codes not zero: {failures}")
        return

    last = outputs[-1]
    if last.get("exit_code") != 0:
        raise AssertionError(
            f"Final CLI call failed (exit_code={last.get('exit_code')}): "
            f"{last.get('stdout', '')[:200]}"
        )

    failures = [output for output in outputs if output.get("exit_code") != 0]
    if len(failures) > len(outputs) * 0.8:
        raise AssertionError(f"Too many CLI failures: {len(failures)}/{len(outputs)}")


def assert_cli_args_contain(result: Any, token: str) -> None:
    for call in _cli_tool_calls(result):
        args = call.arguments.get("args", "")
        if token in args:
            return

    raise AssertionError(f"Expected CLI args to include '{token}', but none did.")


def _resolve_mcp_command() -> list[str]:
    env_command = os.environ.get("EXCEL_MCP_SERVER_COMMAND")
    if env_command:
        import shlex

        return shlex.split(env_command)

    exe_path = REPO_ROOT / "src/ExcelMcp.McpServer/bin/Release/net10.0-windows/Sbroenne.ExcelMcp.McpServer.exe"
    if exe_path.exists():
        return [str(exe_path)]

    project_path = REPO_ROOT / "src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj"
    return [
        "dotnet",
        "run",
        "--project",
        str(project_path),
        "-c",
        "Release",
        "--no-build",
    ]


def _resolve_cli_command() -> str:
    env_command = os.environ.get("EXCEL_CLI_COMMAND")
    if env_command:
        return env_command

    exe_path = REPO_ROOT / "src/ExcelMcp.CLI/bin/Release/net10.0-windows/excelcli.exe"
    if exe_path.exists():
        return str(exe_path)

    return "excelcli"


def _stdio_server(command: str, args: list[str], *, cwd: str | None = None, env: dict[str, str] | None = None) -> dict[str, Any]:
    return {
        "type": "stdio",
        "command": command,
        "args": args,
        "cwd": cwd,
        "env": env or {},
        "tools": ["*"],
    }


@pytest.fixture(scope="session")
def excel_mcp_servers() -> dict[str, Any]:
    command = _resolve_mcp_command()
    return {
        "excel-mcp": _stdio_server(
            command[0],
            command[1:],
            cwd=str(REPO_ROOT),
        )
    }


@pytest.fixture
def excel_cli_servers() -> Any:
    wrapper = TESTS_DIR / "cli_mcp_server.py"
    command = _resolve_cli_command()
    temp_dir = Path(os.environ.get("TEMP", tempfile.gettempdir()))

    pipe = f"llm-{uuid.uuid4().hex}"
    servers = {
        "excel-cli": _stdio_server(
            sys.executable,
            [
                str(wrapper),
                "--command",
                command,
                "--tool-prefix",
                "excel",
                "--timeout",
                "120",
                "--shell",
                "none",
                "--cwd",
                str(temp_dir),
                "--description",
                "Run the Excel CLI with an argument string. Use --help to discover commands.",
            ],
            cwd=str(temp_dir),
            env={"EXCELMCP_CLI_PIPE": pipe},
        )
    }
    yield servers
    environment = {**os.environ, "EXCELMCP_CLI_PIPE": pipe}
    status = subprocess.run(
        [command, "-q", "service", "status"], env=environment,
        capture_output=True, text=True, timeout=30, check=False,
    )
    if status.returncode != 0:
        pytest.fail(f"Owned CLI service status failed: {status.stderr or status.stdout}")
    if json.loads(status.stdout)["running"]:
        stopped = subprocess.run([command, "-q", "service", "stop"], env=environment,
                                 capture_output=True, text=True, timeout=30, check=False)
        if stopped.returncode != 0:
            pytest.fail(f"Owned CLI service cleanup failed: {stopped.stderr or stopped.stdout}")


@pytest.fixture(scope="session")
def excel_mcp_skill_dir() -> str:
    skill = REPO_ROOT / "artifacts" / "generated-skills" / "excel-mcp-report-formatting"
    if not (skill / "SKILL.md").is_file():
        pytest.fail("Generate skills first: pwsh scripts\\Build-AgentSkills.ps1 -GenerateOnly")
    return str(skill.resolve())


@pytest.fixture(scope="session")
def excel_cli_skill_dir() -> str:
    skill = REPO_ROOT / "artifacts" / "generated-skills" / "excel-cli-report-formatting"
    if not (skill / "SKILL.md").is_file():
        pytest.fail("Generate skills first: pwsh scripts\\Build-AgentSkills.ps1 -GenerateOnly")
    return str(skill.resolve())


def build_excel_mcp_eval(
    name: str,
    *,
    servers: dict[str, Any],
    skill_dir: str | None = None,
    allowed_tools: list[str] | None = None,
    instructions: str | None = None,
    model: str = DEFAULT_MODEL,
    max_turns: int = DEFAULT_MAX_TURNS,
    timeout_s: float = DEFAULT_TIMEOUT_S,
    working_directory: str | None = None,
) -> CopilotEval:
    return _build_eval(
        name, servers, skill_dir, allowed_tools, instructions or _MCP_INSTRUCTIONS,
        model, max_turns, timeout_s, working_directory,
    )


def build_excel_cli_eval(
    name: str,
    *,
    servers: dict[str, Any],
    skill_dir: str | None = None,
    allowed_tools: list[str] | None = None,
    instructions: str | None = None,
    model: str = DEFAULT_MODEL,
    max_turns: int = DEFAULT_MAX_TURNS,
    timeout_s: float = DEFAULT_TIMEOUT_S,
    working_directory: str | None = None,
) -> CopilotEval:
    return _build_eval(
        name, servers, skill_dir, allowed_tools, instructions or _CLI_INSTRUCTIONS,
        model, max_turns, timeout_s, working_directory,
    )


def workspace_tool(directory: Path, skill_dir: str | None) -> Tool:
    roots = [directory.resolve()]
    if skill_dir:
        roots.append(Path(skill_dir).resolve())

    def execute(invocation: ToolInvocation) -> ToolResult:
        args = invocation.arguments or {}
        path = (directory / args.get("path", ".")).resolve()
        if not any(path.is_relative_to(root) for root in roots):
            return ToolResult(result_type="denied", text_result_for_llm="Path is outside the supplied task and skill.")
        action = args["action"]
        try:
            if action == "read":
                return ToolResult(text_result_for_llm=path.read_text(encoding="utf-8"))
            if action == "list":
                return ToolResult(text_result_for_llm="\n".join(sorted(p.name for p in path.iterdir())))
            if not path.is_relative_to(roots[0]) or path.suffix.lower() not in {".json", ".csv", ".m", ".txt"}:
                return ToolResult(result_type="denied", text_result_for_llm="Write only task JSON, CSV, M, or text inputs.")
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(args["content"], encoding="utf-8")
            return ToolResult(text_result_for_llm=f"Wrote {path.name}")
        except (OSError, UnicodeError) as error:
            return ToolResult(result_type="failure", text_result_for_llm=str(error))

    return Tool(
        name="workspace",
        description="Read/list supplied task files and skill references, or write JSON, CSV, M, and text inputs in the task folder.",
        parameters={"type": "object", "properties": {
            "action": {"type": "string", "enum": ["read", "list", "write"]},
            "path": {"type": "string"}, "content": {"type": "string"},
        }, "required": ["action", "path"]},
        handler=execute,
        defer="never",
    )


def _build_eval(
    name: str, servers: dict[str, Any], skill_dir: str | None, allowed_tools: list[str] | None,
    instructions: str, model: str, max_turns: int, timeout_s: float,
    working_directory: str | None,
) -> CopilotEval:
    directory = Path(working_directory) if working_directory else TEST_RESULTS_DIR / f"{name}-{uuid.uuid4().hex}"
    directory.mkdir(parents=True, exist_ok=True)
    tools = allowed_tools or ["mcp:*", "builtin:skill", "builtin:ask_user", "builtin:tool_search_tool", "custom:workspace"]
    return CopilotEval(
        name=name,
        model=model,
        client_mode="empty",
        instructions=instructions,
        working_directory=str(directory.resolve()),
        allowed_tools=tools,
        max_turns=max_turns,
        max_tool_calls=DEFAULT_MAX_TOOL_CALLS,
        timeout_s=timeout_s,
        mcp_servers=servers,
        skill_directories=[skill_dir] if skill_dir else [],
        extra_config={"tools": [workspace_tool(directory, skill_dir)], "enable_skills": True},
    )


@pytest.fixture(scope="session")
def fixtures_dir() -> Path:
    return FIXTURES_DIR


@pytest.fixture(scope="session")
def results_dir() -> Path:
    return TEST_RESULTS_DIR
