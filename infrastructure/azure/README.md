# Azure Infrastructure

This directory contains the Azure Application Insights infrastructure used for
telemetry development and validation.

## Files

| File | Purpose |
|---|---|
| `appinsights.bicep` | Application Insights deployment entry point |
| `appinsights.parameters.json` | Production telemetry deployment defaults |
| `appinsights-test.bicep` | Test telemetry deployment entry point |
| `appinsights-test.parameters.json` | Test telemetry deployment defaults |
| `appinsights-resources.bicep` | Shared Application Insights resources and ingestion-time privacy transforms |
| `deploy-appinsights.ps1` | Application Insights deployment helper |
| `configure-analytics-oidc.ps1` | Read-only GitHub Actions workload identity setup |

The workspace transform drops noisy runtime metrics and rejects every
`AppExceptions` row that was not explicitly sanitized by ExcelMcp. Exception
messages and stack traces are never retained; only approved type, source, and
project-owned failure-site classifications are ingested.

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
