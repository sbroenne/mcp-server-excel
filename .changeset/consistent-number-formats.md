---
"excelmcp": patch
---

**Consistent number formats across regional settings**: Range formatting, chart
axes, and PivotTable value fields now use the same US-style format codes through
both the CLI and MCP Server, without corrupting decimal places. Excel continues
to display numbers using the user's regional separators.

**Clean screenshot edges**: Range captures now account for Excel's per-cell pixel
rounding, preventing blank strips at some window sizes and zoom levels.
