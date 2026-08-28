---
"excelmcp": minor
---

**Range writes now accept explicit ISO date values** (#835): use typed `date`, `datetime`, or `datetime-offset` objects inside the existing `values` array. Excel serials honor the workbook date system, timezone-bearing values normalize to UTC, and plain strings remain text.
