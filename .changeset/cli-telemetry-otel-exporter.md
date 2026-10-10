---
"excelmcp": patch
---

**`excelcli` quick commands no longer wait about 2 seconds for telemetry**: the CLI now sends its usage telemetry through OpenTelemetry and the Azure Monitor exporter instead of the Application Insights `TelemetryClient`, whose start-up made a blocking 2-second Azure VM metadata lookup. The event and request records are unchanged, except that `ai.internal.sdkVersion` now ends in `:ext1.8.3`. Setting `OTEL_SDK_DISABLED=true` now turns CLI telemetry off.
