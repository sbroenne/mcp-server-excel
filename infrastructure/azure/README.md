# Azure Infrastructure

This directory contains the Azure Application Insights infrastructure used for
production telemetry. Contributor telemetry testing uses in-memory tests and an
optional local dashboard; see
[telemetry development](../../docs/DEVELOPMENT.md#local-development-with-telemetry).

## Files

| File | Purpose |
|---|---|
| `appinsights.bicep` | Application Insights deployment entry point |
| `appinsights.parameters.json` | Production telemetry deployment defaults |
| `appinsights-resources.bicep` | Shared Application Insights resources and ingestion-time privacy transforms |
| `deploy-appinsights.ps1` | Application Insights deployment helper |
| `configure-analytics-oidc.ps1` | Read-only GitHub Actions workload identity setup |

The workspace transform drops metrics, performance counters, framework traces,
and raw `AppExceptions` rows. Sanitized failures use `UnhandledException` custom
events in `AppEvents`; the transform requires their explicit `Sanitized=true`
marker. Only approved type, source, and project-owned failure-site classifications
are recorded, never exception messages or stacks. Native OpenTelemetry operation
traces map to `AppRequests`, and safe identity fields are in `Properties`.
Redeploy the template when updating the reporting contract; existing clients
remain subject to these fail-closed ingestion rules.

## Public usage analytics

The weekly analytics workflow uses Azure workload identity federation rather
than a client secret. A maintainer with permission to create Entra applications
and role assignments runs:

```powershell
.\infrastructure\azure\configure-analytics-oidc.ps1
```

The script grants the workflow only `Log Analytics Reader` on
`excelmcp-logs` and configures the repository's non-secret Azure identifiers.
It cannot write telemetry or change Azure resources after setup.

## Development and Excel checks

Copilot cloud coding tasks use GitHub-hosted Windows machines without desktop
Excel. Code review uses its separate hosted Linux setup. Run Excel-dependent
checks locally on Windows with Excel installed; see the
[testing guide](../../tests/README.md). Hosted checks do not establish Excel
behavior.
