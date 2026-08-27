---
"Sbroenne.ExcelMcp.McpServer": patch
---

**Prevent exception details from entering telemetry.** Crash analytics now keep
only safe error classifications; exception messages and stack traces are
discarded both in the MCP Server and by an Azure ingestion privacy filter.
