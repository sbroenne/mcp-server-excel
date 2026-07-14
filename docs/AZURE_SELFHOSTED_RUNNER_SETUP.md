# Azure Self-Hosted Runner for Excel Integration Tests

ExcelMcp's integration tests require a real Windows desktop with Microsoft Excel.
The repository uses a cost-optimized Azure VM as a GitHub Actions self-hosted runner.

## Design

| Resource | Configuration | Cost behavior |
|---|---|---|
| VM | `Standard_B2as_v2` (2 vCPU, 8 GB), Windows 11 Pro 24H2 | Started only for test runs |
| OS disk | 128 GB `StandardSSD_LRS` | Billed while the VM is deallocated |
| Public IP | Standard static IPv4 | Billed while allocated |
| Network | VNet, NIC, NSG | RDP allowed only from the configured admin CIDR |
| Auto-shutdown | DevTest Labs schedule | Three-hour watchdog set by each workflow run |

There is no Azure Bastion or NAT Gateway. The VM's public IP provides both outbound
GitHub access and direct RDP, which is the lowest-cost configuration. The expected
cost is approximately $12-15/month for one nightly run; current Azure prices and
actual test duration determine the final amount.

The Windows client image is for development and testing under an eligible Visual
Studio subscription. Do not deploy this template to a subscription without Windows
client dev/test rights.

## Security Model

- GitHub authenticates to Azure using OIDC; there is no Azure client secret.
- The OIDC service principal has `Contributor` only on `rg-excel-runner`.
- RDP port 3389 accepts traffic only from `rdpSourceAddressPrefix`.
- The GitHub runner service runs as the same local Windows account that activates
  Office, so Excel sees the correct user profile and license.
- This runner belongs only to this repository. Do not expose it to workflows from
  untrusted forks.

## Provision Infrastructure

Prerequisites:

- Azure CLI authenticated to the target subscription.
- A strong VM administrator password.
- Your current public IPv4 address.

```powershell
$resourceGroup = "rg-excel-runner"
$location = "eastus2"
$adminIp = (Invoke-RestMethod "https://api.ipify.org?format=json").ip

az group create `
  --name $resourceGroup `
  --location $location

az deployment group create `
  --name excel-runner `
  --resource-group $resourceGroup `
  --template-file infrastructure\azure\azure-runner.bicep `
  --parameters infrastructure\azure\azure-runner.parameters.json `
  adminPassword="<strong-password>" `
  rdpSourceAddressPrefix="$adminIp/32"
```

The template creates the VM, networking, Standard SSD, public IP, NSG, and the
auto-shutdown backstop. It does not install Office or register the GitHub runner.

## Install and Activate Office

1. RDP to the template's `publicIp` output as `azureuser`.
2. Install Office with Microsoft Excel.
3. Open Excel and sign in with the licensed account.
4. Dismiss first-run, privacy, update, and file-association dialogs.
5. In Excel Trust Center, enable **Trust access to the VBA project object model**.
6. Close Excel and reboot once after Office finishes updating.

The VBA Trust Center setting is required by the repository's VBA smoke tests.

## Install the GitHub Runner

Generate a repository registration token. It expires after one hour:

```powershell
$runnerToken = gh api `
  --method POST `
  repos/sbroenne/mcp-server-excel/actions/runners/registration-token `
  --jq ".token"
```

Run `infrastructure\azure\setup-runner.ps1` on the VM from an elevated PowerShell
window. Supply the local account that activated Office:

```powershell
.\setup-runner.ps1 `
  -GithubRepoUrl "https://github.com/sbroenne/mcp-server-excel" `
  -GithubRunnerToken $runnerToken `
  -WindowsServiceAccount ".\azureuser" `
  -WindowsServicePassword "<vm-admin-password>"
```

The script:

1. Installs .NET 10 if necessary.
2. Installs Git for Windows and PowerShell 7 if necessary.
3. Creates the Windows service-profile Desktop folders required by Office COM.
4. Resolves and installs the latest GitHub Actions runner release.
5. Registers labels `self-hosted`, `Windows`, `X64`, and `excel`.
6. Installs the runner as an automatic Windows service under `azureuser`.

The registration token and Windows password are not written to
`C:\runner-setup.log`.

## Configure GitHub OIDC

The workflow reads these repository Actions secrets:

- `AZURE_CLIENT_ID`
- `AZURE_TENANT_ID`
- `AZURE_SUBSCRIPTION_ID`

The Entra app requires this federated credential:

| Field | Value |
|---|---|
| Issuer | `https://token.actions.githubusercontent.com` |
| Subject | `repo:sbroenne/mcp-server-excel:ref:refs/heads/main` |
| Audience | `api://AzureADTokenExchange` |

Assign the app's service principal `Contributor` at this scope only:

```text
/subscriptions/<subscription-id>/resourceGroups/rg-excel-runner
```

## Workflow Lifecycle

`.github/workflows/integration-tests.yml` runs nightly and on manual dispatch:

1. A GitHub-hosted job starts the VM and sets auto-shutdown to three hours from now.
2. The self-hosted job runs the integration projects sequentially with explicit
   hang timeouts, then runs the OnDemand session tests.
3. A GitHub-hosted `always()` job deallocates the VM.
4. The auto-shutdown schedule limits compute cost if the runner never accepts the
   queued job or final cleanup cannot run.

The workflow uploads TRX results for 14 days.

## Operations

Start the VM for maintenance:

```powershell
az vm start --resource-group rg-excel-runner --name vm-excel-runner
```

Stop compute billing after maintenance:

```powershell
az vm deallocate --resource-group rg-excel-runner --name vm-excel-runner
```

Update the RDP source IP:

```powershell
$adminIp = (Invoke-RestMethod "https://api.ipify.org?format=json").ip
az network nsg rule update `
  --resource-group rg-excel-runner `
  --nsg-name vm-excel-runner-nsg `
  --name AllowRdpFromAdmin `
  --source-address-prefixes "$adminIp/32"
```

Check runner state:

```powershell
gh api repos/sbroenne/mcp-server-excel/actions/runners `
  --jq ".runners[] | {name,status,busy,labels:[.labels[].name]}"
```

Delete all billable resources:

```powershell
az group delete --name rg-excel-runner --yes
```

Remove the runner from GitHub before deleting the VM, or remove the stale runner
entry in **Settings > Actions > Runners** afterward.
