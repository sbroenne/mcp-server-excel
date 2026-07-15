# Azure Self-Hosted Runner for Excel Integration Tests

ExcelMcp's integration tests require a real Windows desktop with Microsoft Excel.
The repository uses a cost-optimized Azure VM as a GitHub Actions self-hosted runner.

## Design

| Resource | Configuration | Cost behavior |
|---|---|---|
| VM | `Standard_D2s_v5` (2 Intel vCPU, 8 GB), Windows 11 Pro 24H2 | Started only for test runs |
| OS disk | 128 GB `StandardSSD_LRS` | Billed while the VM is deallocated |
| Public IP | Standard static IPv4 | Billed while allocated |
| Network | VNet, NIC, NSG | RDP allowed only from the configured admin CIDR |
| Auto-shutdown | DevTest Labs schedule | Five-hour watchdog set by each workflow run |

There is no Azure Bastion or NAT Gateway. The VM's public IP provides both outbound
GitHub access and direct RDP, which is the lowest-cost configuration. The expected
cost is approximately $30-40/month for one nightly run; current Azure prices and
actual test duration determine the final amount.

The runner intentionally uses the v5 Intel series. The v6 and v7 D-series sizes
available in East US 2 require an NVMe boot controller, while this installed
Office runner uses SCSI. Moving to those series requires rebuilding the runner
from an NVMe-native image.

The Windows client image is for development and testing under an eligible Visual
Studio subscription. Do not deploy this template to a subscription without Windows
client dev/test rights.

## Security Model

- GitHub authenticates to Azure using OIDC; there is no Azure client secret.
- The OIDC service principal has `Contributor` only on `rg-excel-runner`.
- RDP port 3389 accepts traffic only from `rdpSourceAddressPrefix`.
- The GitHub runner starts in the interactive `azureuser` console session so Excel
  window and clipboard operations have a real desktop.
- Microsoft Sysinternals Autologon stores the Windows password as an LSA secret;
  administrators on the VM can still retrieve it.
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
  -WindowsAccount ".\azureuser" `
  -WindowsPassword "<vm-admin-password>"
```

The script:

1. Installs .NET 10 if necessary.
2. Installs Git for Windows and PowerShell 7 if necessary.
3. Sets the system and `azureuser` locale to `en-US` for deterministic Excel formats.
4. Creates the Windows service-profile Desktop folders required by Office COM.
5. Resolves and installs the latest GitHub Actions runner release.
6. Registers labels `self-hosted`, `Windows`, `X64`, and `excel`.
7. Configures secure automatic logon with Microsoft Sysinternals Autologon.
8. Starts `run.cmd` at interactive user logon instead of as a Windows service.

The registration token and Windows password are not written to `C:\runner-setup.log`.
Reboot after setup so the locale, automatic logon, and interactive runner activate.

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

1. A GitHub-hosted job starts the VM and sets auto-shutdown to five hours from now.
2. The self-hosted job normalizes the runner profile to `en-US`, verifies Excel
   reports `.` as its decimal separator and `,` as its thousands separator, then
   runs the integration projects sequentially with explicit hang timeouts and
   the OnDemand session tests.
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
