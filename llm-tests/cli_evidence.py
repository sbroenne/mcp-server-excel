"""Decode SDK tool evidence and correlate CLI batch results with supplied inputs."""

from __future__ import annotations

import json
import shlex
from pathlib import Path
from typing import Any

from pytest_skill_engineering.copilot.result import CopilotResult, ToolCall


class BatchEvidenceError(AssertionError):
    """The captured batch cannot establish which commands actually ran."""


def _require_batch_evidence(condition: bool, message: str) -> None:
    if not condition:
        raise BatchEvidenceError(message)


def tool_output(call: ToolCall) -> dict[str, Any]:
    text = (call.result or "").strip()
    payload, end = json.JSONDecoder().raw_decode(text)
    if isinstance(payload, str):
        payload = json.loads(payload)
    assert isinstance(payload, dict), f"Tool evidence must be an object, not {type(payload).__name__}"
    trailing = text[end:].strip()
    if trailing:
        duplicate = json.loads(trailing)
        if isinstance(duplicate, dict) and set(duplicate) == {"result"}:
            duplicate = duplicate["result"]
        if isinstance(duplicate, str):
            duplicate = json.loads(duplicate)
        assert duplicate == payload, "Text and structured execution evidence disagree"
    return payload


def cli_args(call: ToolCall) -> list[str]:
    args = shlex.split(call.arguments.get("args", ""))
    return args[1:] if args[:1] == ["-q"] else args


def batch_steps(result: CopilotResult, call: ToolCall) -> list[dict[str, Any]]:
    args = cli_args(call)
    _require_batch_evidence(args[:1] == ["batch"], "Batch evidence requires a batch call")
    payload = tool_output(call)
    lines = payload.get("stdout", "").splitlines()
    exit_code = payload.get("exit_code")
    if type(exit_code) is int and exit_code != 0 and len(lines) == 1:
        failure = json.loads(lines[0])
        if (isinstance(failure, dict) and failure.get("success") is False
                and failure.get("exceptionType") == "CommandParseException"
                and "index" not in failure and "commandIndex" not in failure):
            return []
    input_path = None
    for index, argument in enumerate(args):
        if argument in ("--input", "-i") and index + 1 < len(args):
            input_path = args[index + 1]
        elif argument.startswith("--input="):
            input_path = argument.split("=", 1)[1]
    _require_batch_evidence(bool(input_path) and input_path != "-", "Batch input was not captured")
    base = Path(result.agent.working_directory) if result.agent else Path.cwd()

    def key(path: str) -> str:
        return str((base / path).resolve()).casefold()

    content = None
    for previous in result.all_tool_calls:
        if previous is call:
            break
        if (previous.name == "workspace" and previous.success is True
                and previous.arguments.get("action") == "write"
                and key(previous.arguments["path"]) == key(input_path)):
            content = previous.arguments["content"]
    _require_batch_evidence(content is not None, "Executed batch input must have captured workspace write evidence")
    commands = (json.loads(content) if content.lstrip().startswith("[")
                else [json.loads(line) for line in content.splitlines() if line.strip()])
    _require_batch_evidence(isinstance(commands, list), "Batch inputs must be command rows")
    steps = []
    for line in lines:
        output = json.loads(line)
        _require_batch_evidence(isinstance(output, dict) and isinstance(output.get("index"), int), "Invalid batch output")
        _require_batch_evidence(0 <= output["index"] < len(commands), "Batch result does not match its input")
        command = commands[output["index"]]
        _require_batch_evidence(command["command"].casefold() == output["command"].casefold(), "Batch command identity changed")
        steps.append({**output, "args": command.get("args", {})})
    return steps
