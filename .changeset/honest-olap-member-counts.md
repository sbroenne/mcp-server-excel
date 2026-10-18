---
"excelmcp": patch
---

External OLAP member searches now pass the provider's level count through as `totalCount` and no longer return a calculated `omittedCount`, which could be wrong when the provider's count is approximate. Use the continuation token to know whether more members remain, in both MCP and the CLI.
