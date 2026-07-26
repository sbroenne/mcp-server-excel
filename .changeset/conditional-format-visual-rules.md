---
"excelmcp": minor
---

**Conditional formatting: full support for visual rule types** (#743). `add-rule` can now create colorScale, dataBar, iconSet, top10, aboveAverage, timePeriod, uniqueValues and blanksCondition rules via discrete, LLM-friendly parameters (e.g. `colorScaleMinColor`, `dataBarDirection`, `iconSetId`, `rank`, `aboveBelow`, `datePeriod`). `list-rules` and `list-worksheet-rules` now report each visual rule's type-specific configuration — color-scale stops, data-bar settings, icon-set thresholds, top/bottom, above/below and date period — with colors as `#RRGGBB`, so visual rules can be fully inspected and round-tripped.
