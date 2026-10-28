---
"excelmcp": patch
---

Let Excel determine which files and web addresses it can open, with URL parameters passed unchanged. Remove custom host, file-format, exclusive-access, and visibility checks from both CLI and MCP opening. Excel can open readable files without requiring write access first. Fix finding and reconnecting to personal OneDrive workbooks when Excel reports their web address instead of their synced Windows path.

Keep files selected by different values after `?` separate without showing credentials or those values in reported addresses. Use consistent Windows paths after a failed copy or move between files, so Save As no longer leaves the old path marked as in use.
