---
"excelmcp": minor
---

Replace mode-only calculation actions with native settings read/write, iteration
controls, and explicit workbook precision-as-displayed permission. Add full
recalculation and dependency rebuild. Rename the former workbook scope to
application to accurately describe all open workbooks in the owned Excel process.
Report actual native calculation state without assuming unavailable state is done.
