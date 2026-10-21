---
"excelmcp": patch
---

Add actions to find and attach to workbooks already open in Excel without reopening them. Discovery includes workbooks in Excel processes already managed by this host and identifies workbooks with existing sessions. Newly attached workbooks keep their AutoSave setting, report it to the user, and remain open when the service shuts down. Shutdown closes only workbooks ExcelMcp opened; it leaves Excel running when other workbooks remain and never automatically terminates Excel when the host exits.
