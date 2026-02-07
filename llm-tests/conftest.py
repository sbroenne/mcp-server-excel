"""Fixtures and helpers for ExcelMcp LLM integration tests."""

from __future__ import annotations

import json
import os
import re
import shlex
import tempfile
import uuid
from pathlib import Path
from typing import Any, Iterable

import pytest

from pytest_aitest import Agent, CLIServer, MCPServer, Provider, Skill, Wait

TESTS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TESTS_DIR.parent
FIXTURES_DIR = TESTS_DIR / "Fixtures"
TEST_RESULTS_DIR = TESTS_DIR / "TestResults"
TEST_RESULTS_DIR.mkdir(parents=True, exist_ok=True)

# Skill references are now copied at build time by MSBuild targets in CLI/MCP csproj files.
# Run 'dotnet build -c Release' to update skill references.

# LiteLLM expects AZURE_API_BASE while Azure SDK uses AZURE_OPENAI_ENDPOINT.
if os.environ.get("AZURE_OPENAI_ENDPOINT") and not os.environ.get("AZURE_API_BASE"):
    os.environ["AZURE_API_BASE"] = os.environ["AZURE_OPENAI_ENDPOINT"]

DEFAULT_MODEL = "gpt-4.1"
DEFAULT_RPM = 10
DEFAULT_TPM = 10000
DEFAULT_MAX_TURNS = 20


def pytest_configure(config: pytest.Config) -> None:
    azure_base = os.environ.get("AZURE_API_BASE") or os.environ.get("AZURE_OPENAI_ENDPOINT")
    if azure_base:
        config.option.llm_model = "azure/gpt-4.1"
        config.option.llm_api_base = azure_base


def unique_path(prefix: str, suffix: str = ".xlsx") -> str:
    temp_dir = Path(os.environ.get("TEMP", tempfile.gettempdir()))
    path = temp_dir / f"{prefix}-{uuid.uuid4()}{suffix}"
    return path.as_posix()


def unique_results_path(prefix: str, suffix: str = ".xlsx") -> str:
    path = TEST_RESULTS_DIR / f"{prefix}-{uuid.uuid4()}{suffix}"
    return path.as_posix()


def assert_regex(text: str, pattern: str) -> None:
    if not re.search(pattern, text, re.IGNORECASE | re.MULTILINE):
        raise AssertionError(f"Pattern not found: {pattern}\nText:\n{text}")


def _parse_cli_results(result: Any) -> list[dict[str, Any]]:
    calls = result.tool_calls_for("excel_execute")
    outputs: list[dict[str, Any]] = []
    for call in calls:
        if call.result:
            try:
                outputs.append(json.loads(call.result))
            except json.JSONDecodeError:
                outputs.append({"exit_code": -1, "stdout": call.result, "stderr": ""})
    return outputs


def assert_cli_exit_codes(result: Any) -> None:
    outputs = _parse_cli_results(result)
    if not outputs:
        raise AssertionError("No CLI executions recorded")
    for output in outputs:
        if output.get("exit_code") != 0:
            raise AssertionError(f"CLI exit code not zero: {output}")


def assert_cli_args_contain(result: Any, token: str) -> None:
    calls = result.tool_calls_for("excel_execute")
    for call in calls:
        args = call.arguments.get("args", "")
        if token in args:
            return
    raise AssertionError(f"Expected CLI args to include '{token}', but none did.")


def _resolve_mcp_command() -> list[str]:
    env_command = os.environ.get("EXCEL_MCP_SERVER_COMMAND")
    if env_command:
        return shlex.split(env_command)

    # Windows-specific build with COM interop support
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

    # Windows-specific build with COM interop support
    exe_path = REPO_ROOT / "src/ExcelMcp.CLI/bin/Release/net10.0-windows/excelcli.exe"
    if exe_path.exists():
        return str(exe_path)

    # Fallback to excelcli in PATH
    return "excelcli"


@pytest.fixture(scope="session")
def excel_mcp_server() -> MCPServer:
    return MCPServer(
        command=_resolve_mcp_command(),
        wait=Wait.ready(timeout_ms=30000),
    )


@pytest.fixture(scope="session")
def excel_cli_server() -> CLIServer:
    command = _resolve_cli_command()
    temp_dir = Path(os.environ.get("TEMP", tempfile.gettempdir()))
    return CLIServer(
        name="excel-cli",
        command=command,
        tool_prefix="excel",
        shell="none",
        cwd=str(temp_dir),
        discover_help=False,  # Skill Rule 0 requires LLM to run --help first
        description="Excel CLI automation. Run 'excelcli --help' to discover available commands before use.",
    )


@pytest.fixture(scope="session")
def excel_mcp_skill() -> Skill:
    return Skill.from_path(REPO_ROOT / "skills/excel-mcp")


@pytest.fixture(scope="session")
def excel_cli_skill() -> Skill:
    return Skill.from_path(REPO_ROOT / "skills/excel-cli")


@pytest.fixture(scope="session")
def fixtures_dir() -> Path:
    return FIXTURES_DIR


@pytest.fixture(scope="session")
def results_dir() -> Path:
    return TEST_RESULTS_DIR
