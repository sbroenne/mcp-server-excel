---
"excelmcp": patch
---

Let Excel determine which files and web addresses it can open, with URL parameters passed unchanged. Remove custom host, file-format, exclusive-access, and visibility checks from both CLI and MCP opening. Excel can open readable files without requiring write access first. Fix finding and reconnecting to personal OneDrive workbooks when Excel reports their web address instead of their synced Windows path.
