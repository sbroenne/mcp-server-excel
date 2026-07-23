# Azure Excel Integration Runner

This directory contains the cost-optimized Azure infrastructure and setup script
for the repository's Windows + Excel GitHub Actions runner.

## Files

| File | Purpose |
|---|---|
| `azure-runner.bicep` | VM, Standard SSD, public IP, NSG, network, auto-shutdown |
| `azure-runner.parameters.json` | Non-secret deployment defaults |
| `setup-runner.ps1` | Installs prerequisites and registers the latest runner for a selected local Windows account |

The VM is deallocated outside test and maintenance windows. Direct RDP is restricted
to one administrator CIDR; Azure Bastion and NAT Gateway are intentionally omitted
to minimize standing cost.

See [Azure Self-Hosted Runner Setup](../../docs/AZURE_SELFHOSTED_RUNNER_SETUP.md)
for provisioning, Office activation, OIDC configuration, operations, and cleanup.
