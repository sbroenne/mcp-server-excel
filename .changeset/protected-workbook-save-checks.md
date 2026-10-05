---
"excelmcp": patch
---

Protected workbooks now use the signed-in user's Excel editing permissions instead of always opening read-only. Workbook changes reject genuine read-only access, and saves and Save As report an error when Excel does not accept the save, leaving the session open for inspection.
