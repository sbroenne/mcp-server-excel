---
"excelmcp": patch
---

**Fix `conditionalformat add` throwing on `borderStyle`/`borderColor`** (#737). Writing border formatting on a conditional-format rule threw `COMException: Unable to set the LineStyle property of the Border class`. Root cause: `FormatCondition.Borders` is a 4-item collection indexed 1-4 (left/top/bottom/right), unlike `Range.Borders` which uses the `xlEdgeLeft`/`Top`/`Bottom`/`Right` constants (7-10) — writing (and reading) via those out-of-range indices silently returned an unbound placeholder that threw on write and reported blank values on read. Both the write path (`add`) and the read path (`list-rules`/`list-worksheet-rules`) now use the correct 1-4 indices, so border style and color round-trip correctly.
