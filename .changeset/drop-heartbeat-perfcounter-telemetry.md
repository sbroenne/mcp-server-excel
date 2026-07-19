---
"excelmcp": patch
---

**Further reduced MCP Server telemetry ingestion cost.** The MCP Server no longer reports Application Insights heartbeat (`HeartbeatState`) or performance counter telemetry (`Requests/Sec`, `Private Bytes`, `% Processor Time`, etc.) — these were still being emitted despite existing configuration flags intended to disable them, and accounted for roughly a third of remaining telemetry ingestion volume. A Log Analytics ingestion-time transform (Data Collection Rule) now also drops these rows, plus the `http.client.request.duration` HTTP-client metric, as a reliable server-side backstop independent of in-process SDK behavior. This Data Collection Rule is now defined in `infrastructure/azure/appinsights-resources.bicep` (previously only configured manually in Azure, at risk of being lost on redeployment).
