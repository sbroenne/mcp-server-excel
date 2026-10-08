---
"excelmcp": patch
---

MCP and CLI session listings now distinguish an open Excel dialog from a busy query and explain when to check Excel for a prompt. Detection works without calling Excel, includes dialogs hosted in another process, and never reads credentials or responds to prompts. Save and close remain blocked until Excel is ready.
