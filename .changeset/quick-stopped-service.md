---
"excelmcp": patch
---

`excelcli service status` and `excelcli session list` now report a stopped service promptly instead of waiting two seconds, while continuing to wait for a service that is starting.
