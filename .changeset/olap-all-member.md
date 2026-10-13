---
"excelmcp": patch
---

Fix PivotTable reads and report-filter results incorrectly treating an OLAP All member named `[AllMember]` as a single selected member. Both MCP and CLI now identify single-selection All state from the hierarchy metadata exposed by Excel, regardless of the server's member naming.
