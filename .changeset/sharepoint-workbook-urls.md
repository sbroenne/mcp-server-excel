---
"excelmcp": minor
---

Open existing SharePoint and OneDrive for Business workbooks directly from their HTTPS file URLs in MCP and excelcli, without syncing or downloading a separate copy first. Visible Excel handles sign-in and editing permissions; cloud AutoSave is disabled so explicit save and discard behavior stays consistent with local workbooks.

Workbook information now reports the live AutoSave status through `autoSaveOn`. Older Excel versions without AutoSave report it as disabled and can still open cloud workbooks; unexpected AutoSave access errors remain failures.
