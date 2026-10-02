"""Shared comparison evidence and independently checked completion."""

from __future__ import annotations

import json
import asyncio
import os
from collections import defaultdict
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path, PureWindowsPath
from typing import Any, Callable

from pytest_skill_engineering.copilot.result import CopilotResult

from consent_scenarios import cli_boolean_option
from cli_evidence import BatchEvidenceError, batch_steps, cli_args, tool_output


def check_attempt_budget(planned: int, prior: int, ceiling: int) -> None:
    if not 1 <= ceiling <= 100 or prior < 0 or planned < 0 or planned + prior > ceiling:
        raise ValueError("Planned and prior attempts must fit the approved ceiling (at most 100)")


def assert_explicit_save(result: CopilotResult, transport: str) -> None:
    assert result.evidence_complete, f"Incomplete execution evidence: {result.capture_errors}"
    closes = []
    for call in result.all_tool_calls:
        if transport == "mcp" and call.name == "excel-mcp-file" and call.arguments.get("action") == "close":
            closes.append((call, call.arguments.get("save") is True, tool_output(call)))
        elif transport == "cli" and call.name in {"excel_execute", "excel-cli-excel_execute"}:
            args = cli_args(call)
            if "--help" in args:
                continue
            if args[:2] == ["session", "close"]:
                payload = tool_output(call)
                assert payload["exit_code"] == 0, payload
                closes.append((call, cli_boolean_option(args, "--save"), json.loads(payload["stdout"])))
            elif args[:1] == ["batch"]:
                for step in batch_steps(result, call):
                    if step["command"].casefold() == "session.close":
                        closes.append((call, step["args"].get("save") is True, step))
    assert closes, "The request requires an explicit save and close"
    call, saving, payload = closes[-1]
    assert saving and call.success is True, "The request requires an explicit successful save and close"
    assert payload.get("success") is True, f"Save and close failed: {payload}"


def completed_operations(result: CopilotResult) -> list[str]:
    operations = []
    for call in result.all_tool_calls:
        if call.success is not True:
            continue
        if call.name == "excel-mcp-screenshot":
            operations.append(f"screenshot.{call.arguments.get('action')}")
        elif call.name.startswith("excel-mcp-"):
            payload = tool_output(call)
            if payload.get("success") is True:
                operations.append(f"{call.name.removeprefix('excel-mcp-')}.{call.arguments.get('action')}")
        elif call.name in {"excel_execute", "excel-cli-excel_execute"}:
            args = cli_args(call)
            if not args or "--help" in args:
                continue
            if args[0] == "batch":
                for step in batch_steps(result, call):
                    if step.get("success") is True:
                        operations.append(step["command"])
            elif len(args) >= 2 and tool_output(call).get("exit_code") == 0:
                operations.append(".".join(args[:2]))
    return operations


def execution_metrics(result: CopilotResult) -> dict[str, Any]:
    premium = None
    for event in result.raw_events:
        event_type = getattr(event.type, "value", event.type)
        if event_type == "session.shutdown":
            value = getattr(event.data, "_total_premium_requests", None)
            if value is not None:
                premium = float(value)
    return {
        "tokens": result.total_tokens,
        "premium_requests": premium,
        "tool_calls": len(result.all_tool_calls),
        "skill_reads": [
            {"tool": call.name, "arguments": call.arguments, "success": call.success}
            for call in result.all_tool_calls
            if call.name == "skill" or (
                call.name == "workspace" and call.arguments.get("action") == "read" and (
                    PureWindowsPath(call.arguments.get("path", "")).name.lower() == "skill.md"
                    or "references" in PureWindowsPath(call.arguments.get("path", "")).parts
                )
            )
        ],
        "duration_ms": result.duration_ms,
        "evidence_complete": result.evidence_complete,
        "model_used": result.model_used,
        "skill_discovery": asdict(result.skill_discovery) if result.skill_discovery is not None else None,
        "stop_reason": result.stop_reason,
        "batch_subcommands": _batch_commands(result),
    }


async def record_execution(
    record: dict[str, Any], path: Path, evaluate: Callable, agent: Any,
    prompt: str, verify: Callable[[CopilotResult], dict[str, Any]],
) -> None:
    def save() -> None:
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        temporary = path.with_suffix(".json.tmp")
        with temporary.open("w", encoding="utf-8") as stream:
            json.dump(record, stream, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(path)

    journal_path = path.with_suffix(".events.jsonl")
    previous_event = agent.extra_config.get("on_event") if agent is not None else None

    def on_event(event: Any) -> None:
        with journal_path.open("a", encoding="utf-8") as stream:
            stream.write(event.model_dump_json() + "\n")
            stream.flush()
            os.fsync(stream.fileno())
        record["events_received"] += 1
        save()
        if previous_event is not None:
            previous_event(event)

    record.update(status="running", events_received=0, event_journal=journal_path.name)
    save()
    if agent is not None:
        agent.extra_config["on_event"] = on_event
    try:
        result = await evaluate(agent, prompt)
        record["execution"] = {
            "success": result.success, "error": result.error,
            "capture_errors": result.capture_errors, "final_response": result.final_response,
            "calls": [{"name": call.name, "arguments": call.arguments, "result": call.result,
                       "error": call.error, "success": call.success,
                       "completion_received": call.completion_received}
                      for call in result.all_tool_calls],
        }
        record["status"] = "verifying"
        record.update({**execution_metrics(result), "category": "task", "verification_error": None})
        if "expected_skill_read" in record:
            read_skill = any(read["success"] is True for read in record["skill_reads"])
            record["selection_matches_intent"] = read_skill == record["expected_skill_read"]
        save()
        if result.stop_reason in {"timeout", "tool_budget_exceeded"}:
            record.update(category="budget", verification_error=result.error or result.stop_reason)
        elif not result.evidence_complete:
            record.update(category="capture", verification_error=str(result.capture_errors))
        elif not result.success:
            record.update(category="execution", verification_error=result.error)
        elif record.get("condition") == "with-skill" and not _matching_skill_available(
            result, record.get("skill_name", f"excel-{record['transport']}"),
        ):
            record.update(category="capture", verification_error="Treatment lacks complete, matching session skill discovery")
        else:
            try:
                record["verification"] = verify(result)
                record.update(passed=True, category="verified")
            except BatchEvidenceError as error:
                record.update(category="capture", verification_error=str(error))
            except AssertionError as error:
                record["verification_error"] = str(error) or "Independent workbook check failed"
        record["status"] = "finished"
    except (asyncio.CancelledError, KeyboardInterrupt, SystemExit) as error:
        record.update(status="interrupted", category="harness",
                      verification_error=f"{type(error).__name__}: {error}")
        raise
    except Exception as error:
        record.update(status="finished", category="harness", verification_error=f"{type(error).__name__}: {error}")
        raise
    finally:
        if agent is not None:
            if previous_event is None:
                agent.extra_config.pop("on_event", None)
            else:
                agent.extra_config["on_event"] = previous_event
        save()


def _matching_skill_available(result: CopilotResult, expected_name: str) -> bool:
    discovery = result.skill_discovery
    return discovery is not None and discovery.complete and not discovery.errors and (
        sorted(skill.name for skill in discovery.skills if skill.enabled) == [expected_name]
    )


def _batch_commands(result: CopilotResult) -> int:
    count = 0
    for call in result.all_tool_calls:
        if call.name not in {"excel_execute", "excel-cli-excel_execute"} or not call.result:
            continue
        args = cli_args(call)
        if args[:1] != ["batch"] or "--help" in args:
            continue
        payload = tool_output(call)
        for line in payload.get("stdout", "").splitlines():
            try:
                operation = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(operation, dict) and ("commandIndex" in operation or "index" in operation):
                count += 1
    return count


def summarize(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        groups[(record["transport"], record["task"], record["condition"])].append(record)
    rows = []
    for (transport, task, condition), attempts in sorted(groups.items()):
        verified = sum(record["passed"] is True for record in attempts)
        row = {"transport": transport, "task": task, "condition": condition,
               "attempts": len(attempts), "verified": verified,
               "failures": [record["category"] for record in attempts if not record["passed"]]}
        for field, output in (("tokens", "tokens_per_verified"), ("premium_requests", "requests_per_verified")):
            values = [record.get(field) for record in attempts]
            row[output] = sum(values) / verified if verified and all(value is not None for value in values) else None
        rows.append(row)
    return rows


def paired_results(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    pairs: dict[tuple[str, str, int], dict[str, dict[str, Any]]] = defaultdict(dict)
    for record in records:
        key = (record["transport"], record["task"], record["repetition"])
        condition = record["condition"]
        assert condition not in pairs[key], f"Duplicate comparison condition: {key}, {condition}"
        pairs[key][condition] = record
    rows = []
    for (transport, task, repetition), conditions in sorted(pairs.items()):
        baseline = conditions.get("without-skill")
        treatment = conditions.get("with-skill")
        both = bool(baseline and treatment and baseline["passed"] and treatment["passed"])
        saving = None
        if both and baseline.get("tokens") is not None and treatment.get("tokens") is not None:
            if baseline["tokens"] > 0:
                saving = 100 * (baseline["tokens"] - treatment["tokens"]) / baseline["tokens"]
        rows.append({
            "transport": transport, "task": task, "repetition": repetition,
            "without_skill": baseline, "with_skill": treatment,
            "both_verified": both, "token_saving_percent": saving,
        })
    return rows
