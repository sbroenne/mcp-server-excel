---
"excelmcp": patch
---

**Consistent number formats across regional settings**: Range formatting, chart
axes, and PivotTable value fields now use the same US-style format codes through
both the CLI and MCP Server, without corrupting decimal places or comparison
thresholds in custom formats. Locale-specific date letters inside quoted or
escaped literal text no longer prevent regional decimal translation. Excel continues
to display numbers using the user's regional separators.

**Clean screenshot edges**: Range captures now account for Excel's per-cell pixel
rounding, preventing blank strips at some window sizes and zoom levels.

**Reliable Excel shutdown**: Release unused automation metadata before closing
Excel to avoid stalled cleanup, and report session save and cleanup failures instead of
silently ignoring them. All sessions still receive a teardown attempt when one
auto-save fails. Normal post-Quit cleanup keeps its existing total time
limit and targets only the session's verified Excel process.

**Public command parity**: Range value, formula-read, and clear commands now
accept an empty sheet name when the range address is a workbook named range.
Setting an empty chart title now hides the title as documented.
