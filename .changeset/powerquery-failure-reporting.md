---
"excelmcp": patch
---

Power Query creation and updates now identify privacy-blocked failures correctly and explain which query changes remain in the workbook. MCP and CLI errors also include the inspected query and load state, so callers can recover without accidentally repeating creation or assuming failed updates were undone.
