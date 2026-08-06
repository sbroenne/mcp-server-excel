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
| `appinsights-resources.bicep` | Shared Application Insights resources |
| `deploy-appinsights.ps1` | Application Insights deployment helper |
