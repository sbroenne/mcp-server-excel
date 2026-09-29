---
"excelmcp": patch
---

Preserve unrelated Copilot settings when installing the MCP plugin globally.
Plugin downloads now stop if the shared installation lock cannot be acquired,
and update their cached state without exposing partially written JSON.
The optional CLI installer repairs missing launchers without overwriting their
existing partner and supports installation paths containing apostrophes or
non-ASCII characters.
Failed launcher writes restore the previous installation, or retain recovery
backups and report both errors if automatic restoration is blocked.
