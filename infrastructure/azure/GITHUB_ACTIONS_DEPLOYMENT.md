# GitHub Actions Deployment

The former workflow-driven Bastion deployment has been replaced by a cheaper,
validated Bicep deployment with direct IP-restricted RDP. GitHub Actions uses OIDC
only to start, watchdog, and deallocate the VM around integration-test runs.

Use the canonical
[Azure Self-Hosted Runner Setup](../../docs/AZURE_SELFHOSTED_RUNNER_SETUP.md)
for deployment and maintenance.
