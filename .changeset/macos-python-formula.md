---
"excelmcp": patch
---

Enable Python in Excel formula writes on macOS after exact CLI and MCP Formula2
round-trip acceptance. Keep result reads gated because Excel 16.113.2 rejected
the native Apple Events read through the public MCP entry point.
