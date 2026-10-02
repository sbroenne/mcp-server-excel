---
"Sbroenne.ExcelMcp.McpServer": minor
"Sbroenne.ExcelMcp.CLI": minor
---

Add workbook cell-style listing, complete native inspection, creation from one
stored source cell, and custom-style updates/deletion. Updates can affect all
existing users; omitted inclusion flags are preserved. Built-in styles are
inspectable but read-only. Cell styles expose six native border positions;
inside borders remain a range-formatting operation.

Correct range `get-style` to report Excel's actual built-in/custom status instead
of assuming every registered style is built-in.
