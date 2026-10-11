---
"excelmcp": patch
---

**`excelcli` no longer pauses while telemetry starts and stops**: `--help`, `--version` and the background service skip telemetry setup, and the telemetry library's own health reports, delivery counters and Live Metrics connection are turned off. Usage telemetry is unchanged.
