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

TESTS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TESTS_DIR.parent
FIXTURES_DIR = TESTS_DIR / "Fixtures"
TEST_RESULTS_DIR = TESTS_DIR / "TestResults"
TEST_RESULTS_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_MODEL = "gpt-4.1"
DEFAULT_MAX_TURNS = 20
DEFAULT_MAX_RETRIES = 3
DEFAULT_TIMEOUT_S = 600.0

_MCP_INSTRUCTIONS = (
    "You are an Excel automation assistant. Use the available MCP tools to complete the "
    "workbook task end-to-end. Save workbooks when the task asks for persistence."
)
_CLI_INSTRUCTIONS = (
    "You are an Excel CLI automation assistant. Use the excel CLI tool to complete the "
    "workbook task end-to-end. Run --help when you need command discovery, and prefer "
    "file-based arguments for large JSON or M code payloads."
)


def pytest_collection_modifyitems(items: list[pytest.Item]) -> None:
    for item in items:
        fixturenames = set(getattr(item, "fixturenames", []))
        if "copilot_eval" in fixturenames and not any(m.name == "copilot" for m in item.iter_markers()):
            item.add_marker(pytest.mark.copilot)


def _has_github_auth() -> bool:
    if os.environ.get("GITHUB_TOKEN"):
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


@pytest.fixture(scope="session", autouse=True)
def github_auth() -> None:
    if not _has_github_auth():
        pytest.skip(
            "GitHub auth required for pytest-skill-engineering Copilot tests. "
            "Set GITHUB_TOKEN or run `gh auth login`."
        )


def unique_path(prefix: str, suffix: str = ".xlsx") -> str:
    temp_dir = Path(os.environ.get("TEMP", tempfile.gettempdir()))
    return (temp_dir / f"{prefix}-{uuid.uuid4()}{suffix}").as_posix()


def unique_results_path(prefix: str, suffix: str = ".xlsx") -> str:
    return (TEST_RESULTS_DIR / f"{prefix}-{uuid.uuid4()}{suffix}").as_posix()


def assert_regex(text: str | None, pattern: str) -> None:
    haystack = text or ""
    if not re.search(pattern, haystack, re.IGNORECASE | re.MULTILINE):
        raise AssertionError(f"Pattern not found: {pattern}\nText:\n{haystack}")


def _parse_cli_results(result: Any) -> list[dict[str, Any]]:
    outputs: list[dict[str, Any]] = []
    for call in result.tool_calls_for("excel_execute"):
        payload = call.result or ""
        if not payload:
            continue

        try:
            outputs.append(json.loads(payload))
        except json.JSONDecodeError:
            outputs.append({"exit_code": -1, "stdout": payload, "stderr": ""})

    return outputs


def assert_cli_exit_codes(result: Any, *, strict: bool = False) -> None:
    outputs = _parse_cli_results(result)
    if not outputs:
        raise AssertionError("No CLI executions recorded")

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

    failed = sum(1 for output in outputs if output.get("exit_code") != 0)
    if failed > len(outputs) * 0.8:
        raise AssertionError(f"Too many CLI failures: {failed}/{len(outputs)} calls failed")


def assert_cli_args_contain(result: Any, token: str) -> None:
    for call in result.tool_calls_for("excel_execute"):
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


@pytest.fixture(scope="session")
def excel_cli_servers() -> dict[str, Any]:
    wrapper = TESTS_DIR / "cli_mcp_server.py"
    command = _resolve_cli_command()
    temp_dir = Path(os.environ.get("TEMP", tempfile.gettempdir()))

    return {
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
                "Excel CLI automation. Run 'excelcli --help' to discover available commands before use.",
            ],
            cwd=str(REPO_ROOT),
        )
    }


@pytest.fixture(scope="session")
def excel_mcp_skill_dir() -> str:
    return str((REPO_ROOT / "skills/excel-mcp").resolve())


@pytest.fixture(scope="session")
def excel_cli_skill_dir() -> str:
    return str((REPO_ROOT / "skills/excel-cli").resolve())


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
) -> CopilotEval:
    skill_directories = [skill_dir] if skill_dir else []
    return CopilotEval(
        name=name,
        model=model,
        instructions=instructions or _MCP_INSTRUCTIONS,
        working_directory=str(REPO_ROOT),
        allowed_tools=allowed_tools,
        max_turns=max_turns,
        timeout_s=timeout_s,
        max_retries=DEFAULT_MAX_RETRIES,
        mcp_servers=servers,
        skill_directories=skill_directories,
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
) -> CopilotEval:
    skill_directories = [skill_dir] if skill_dir else []
    return CopilotEval(
        name=name,
        model=model,
        instructions=instructions or _CLI_INSTRUCTIONS,
        working_directory=str(REPO_ROOT),
        allowed_tools=allowed_tools,
        max_turns=max_turns,
        timeout_s=timeout_s,
        max_retries=DEFAULT_MAX_RETRIES,
        mcp_servers=servers,
        skill_directories=skill_directories,
    )


@pytest.fixture(scope="session")
def fixtures_dir() -> Path:
    return FIXTURES_DIR


@pytest.fixture(scope="session")
def results_dir() -> Path:
    return TEST_RESULTS_DIR
