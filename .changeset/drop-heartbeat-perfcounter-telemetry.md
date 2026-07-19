---
"excelmcp": patch
---

**Further reduced Log Analytics ingestion cost for MCP Server telemetry.** Application Insights heartbeat (`HeartbeatState`) and performance counter telemetry (`Requests/Sec`, `Private Bytes`, `% Processor Time`, etc.) are no longer ingested, despite the Application Insights SDK providing no in-process way to disable them in this version — they accounted for roughly a third of remaining telemetry ingestion volume. A Log Analytics ingestion-time transform (Data Collection Rule) now drops these rows server-side, plus the previously-missed `http.client.request.duration` HTTP-client metric. This Data Collection Rule is now defined in `infrastructure/azure/appinsights-resources.bicep` (previously only configured manually in Azure, at risk of being lost on redeployment).
