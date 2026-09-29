---
"excelmcp": patch
---

**Simpler plugin startup:** The Excel MCP and CLI plugins now use the published npm packages through `npx` by default, while retaining the verified GitHub Release downloader as a fallback when Node.js is unavailable.
