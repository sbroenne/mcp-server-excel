"""Small checked subprocess adapter for the supported excelcli entry point."""

import json
from pathlib import Path
import subprocess


class Excel:
    def __init__(self, executable, log_dir):
        self.executable = str(Path(executable).resolve())
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.session = None
        self.sequence = 0

    def invoke(self, *args, timeout=600):
        result = subprocess.run(
            [self.executable, "-q", *map(str, args)],
            capture_output=True, text=True, encoding="utf-8", timeout=timeout,
        )
        if result.returncode:
            raise RuntimeError(f"excelcli {' '.join(map(str, args[:3]))}: {result.stderr}\n{result.stdout}")
        payload = json.loads(result.stdout)
        if payload.get("success") is False:
            raise RuntimeError(json.dumps(payload))
        return payload

    def create(self, path):
        path = Path(path).resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists():
            raise FileExistsError(f"Refusing to overwrite {path}")
        self.session = self.invoke("session", "create", path, "--timeout-seconds", "600", "--show")["sessionId"]
        return self.session

    def open(self, path, show=False):
        options = ["--show"] if show else []
        self.session = self.invoke("session", "open", Path(path).resolve(), "--timeout-seconds", "600", *options)["sessionId"]
        return self.session

    def batch(self, commands):
        self.sequence += 1
        source = self.log_dir / f"batch-{self.sequence:03}.json"
        source.write_text(json.dumps(commands), encoding="utf-8")
        result = subprocess.run(
            [self.executable, "-q", "batch", "--session", self.session,
             "--input", str(source.resolve()), "--stop-on-error"],
            capture_output=True, text=True, encoding="utf-8", timeout=1800,
        )
        (self.log_dir / f"batch-{self.sequence:03}.ndjson").write_text(result.stdout, encoding="utf-8")
        rows = [json.loads(line) for line in result.stdout.splitlines() if line.strip()]
        failed = [row for row in rows if row.get("success") is False]
        if result.returncode or failed or len(rows) != len(commands):
            raise RuntimeError(
                f"Batch {self.sequence} failed ({len(rows)}/{len(commands)} results): "
                f"{result.stderr}\n{json.dumps(failed or rows[-1:], ensure_ascii=False)}"
            )
        return rows

    def call(self, command, **args):
        return self.batch([{"command": command, "args": args}])[0]

    def close(self, save):
        sessions = self.invoke("session", "list")["sessions"]
        owned = next((row for row in sessions if row["sessionId"] == self.session), None)
        if not owned or not owned["canClose"]:
            raise RuntimeError(f"Session {self.session} missing or busy; do not force cleanup")
        flags = ["--save"] if save else []
        result = self.invoke("session", "close", "--session", self.session, *flags)
        self.session = None
        return result
