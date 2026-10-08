---
"excelmcp": patch
---

Keep the workbook and session usable when a refresh starts between the initial readiness check and closing. Retry closing after the refresh completes; any save already completed before the refusal remains saved.
