"""CLI-specific fixtures: auto-start/stop the CLI daemon for LLM tests."""

from __future__ import annotations

import subprocess
import time
from pathlib import Path
from typing import Generator

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent.parent


def _resolve_cli_exe() -> Path:
    """Find the built excelcli.exe."""
    exe = REPO_ROOT / "src/ExcelMcp.CLI/bin/Release/net10.0-windows/excelcli.exe"
    if exe.exists():
        return exe
    raise FileNotFoundError(f"excelcli.exe not found at {exe}. Run: dotnet build -c Release")


@pytest.fixture(scope="session", autouse=True)
def cli_daemon() -> Generator[subprocess.Popen, None, None]:
    """Start the CLI daemon before any CLI LLM test runs, stop it after."""
    exe = _resolve_cli_exe()

    # Start daemon process
    proc = subprocess.Popen(
        [str(exe), "service", "run"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )

    # Wait for daemon to be ready (poll service status)
    for i in range(20):
        try:
            result = subprocess.run(
                [str(exe), "-q", "service", "status"],
                capture_output=True,
                text=True,
                timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            if result.returncode == 0 and '"running":true' in result.stdout:
                break
        except (subprocess.TimeoutExpired, OSError):
            pass
        time.sleep(0.5)
    else:
        proc.kill()
        raise RuntimeError("CLI daemon did not start within 10 seconds")

    yield proc

    # Stop daemon gracefully
    try:
        subprocess.run(
            [str(exe), "-q", "service", "stop"],
            capture_output=True,
            timeout=5,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        proc.wait(timeout=10)
    except (subprocess.TimeoutExpired, OSError):
        proc.kill()
        proc.wait(timeout=5)
