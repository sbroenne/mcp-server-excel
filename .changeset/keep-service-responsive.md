---
"excelmcp": patch
---

Keep the CLI service available when requests pile up behind a busy workbook. Repeated requests to the same busy session are rejected promptly, while health checks and work on other sessions remain available. The service no longer shuts down automatically when idle; use `service stop` to stop it.
