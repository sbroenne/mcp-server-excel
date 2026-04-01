from __future__ import annotations

import argparse
import json
import os
import shlex
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Expose a CLI command as a simple MCP tool.")
    parser.add_argument("--command", required=True)
    parser.add_argument("--tool-prefix", required=True)
    parser.add_argument("--cwd")
    parser.add_argument("--shell")
    parser.add_argument("--timeout", type=int, default=30)
    parser.add_argument("--env-json", default="{}")
    parser.add_argument("--description", default="")
    return parser.parse_args()


def _run_command(command: str, args: str, *, cwd: str | None, shell: str | None, timeout: int, env: dict[str, str]) -> dict[str, Any]:
    start = time.perf_counter()
    full_cmd = command if not args else f"{command} {args}"

    if shell == "none":
        cmd = shlex.split(command, posix=(sys.platform != "win32"))
        if args:
            cmd.extend(shlex.split(args, posix=True))
    elif shell in ("powershell", "pwsh"):
        shell_exe = "powershell" if shell == "powershell" else "pwsh"
        cmd = [shell_exe, "-NoProfile", "-NonInteractive", "-Command", full_cmd]
    elif shell == "cmd":
        cmd = ["cmd", "/C", full_cmd]
    else:
        cmd = [shell or "bash", "-c", full_cmd]

    try:
        completed = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            cwd=cwd,
            env={**os.environ, **env},
            check=False,
        )
        return {
            "command": command,
            "args": args,
            "full_cmd": full_cmd,
            "exit_code": completed.returncode,
            "stdout": completed.stdout,
            "stderr": completed.stderr,
            "duration_ms": int((time.perf_counter() - start) * 1000),
        }
    except subprocess.TimeoutExpired:
        return {
            "command": command,
            "args": args,
            "full_cmd": full_cmd,
            "exit_code": -1,
            "stdout": "",
            "stderr": f"Error: Command timed out after {timeout} seconds",
            "duration_ms": int((time.perf_counter() - start) * 1000),
        }
    except Exception as ex:
        return {
            "command": command,
            "args": args,
            "full_cmd": full_cmd,
            "exit_code": -1,
            "stdout": "",
            "stderr": f"Error: {ex}",
            "duration_ms": int((time.perf_counter() - start) * 1000),
        }


def main() -> None:
    options = _parse_args()
    env = json.loads(options.env_json)
    tool_name = f"{options.tool_prefix}_execute"
    server = FastMCP(f"{options.tool_prefix}-cli")

    @server.tool(
        name=tool_name,
        description=options.description or f"Run {Path(options.command).name} with an args string.",
    )
    def execute(args: str = "") -> str:
        result = _run_command(
            options.command,
            args,
            cwd=options.cwd,
            shell=options.shell,
            timeout=options.timeout,
            env=env,
        )
        return json.dumps(result)

    server.run()


if __name__ == "__main__":
    main()
