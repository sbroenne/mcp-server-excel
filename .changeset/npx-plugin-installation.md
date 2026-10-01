---
"excelmcp": patch
---

Removed the optional global installation scripts from both Copilot plugins. Use the existing npx launch commands without changing PATH or adding a separate global MCP configuration. The CLI plugin retains its Windows launcher for quoted JSON arguments, and installation guidance now matches the npm-based setup.
