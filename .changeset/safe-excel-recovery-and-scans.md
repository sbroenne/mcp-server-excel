---
"excelmcp": patch
---

CLI service status and stop now report Excel copies left by an earlier service without closing them, and those copies no longer block startup. Session status avoids waiting for workbook-location bookkeeping and includes the current operation's start time. Error-cell scans use whole-range checks instead of slow cell-by-cell checks when Excel reports no matching cells.
