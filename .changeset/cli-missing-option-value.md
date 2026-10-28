---
"excelmcp": patch
---

**Clear error for a CLI option with no value**: when a value such as `--session <id>` or `--sheet-name <name>` is left out, `excelcli` now reports "Option '--session' requires a value." It no longer sends the command with the placeholder `__default_command` as the value, which produced confusing "Session '__default_command' not found" errors.
