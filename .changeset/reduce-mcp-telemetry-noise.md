---
"excelmcp": patch
---

**Reduced MCP Server telemetry noise and cost.** The MCP Server no longer reports the .NET runtime's built-in HTTP-client connection-pool metrics (`http.client.open_connections`, `http.client.active_requests`, `http.client.connection.duration`, `http.client.request.time_in_queue`, `http.client.request.duration`) to Application Insights. These were emitted automatically by the telemetry SDK regardless of actual traffic and accounted for the large majority of telemetry ingestion volume, without providing any useful signal for this tool.
