# Azure Self-Hosted Runner Setup for Excel Integration Testing

> **Purpose:** Enable full Excel COM integration testing in CI/CD using Azure-hosted Windows VM with Microsoft Excel

## Quick Navigation

**Choose your path:**

| Scenario | Guide | Time |
|----------|-------|------|
| **🚀 New setup (no VM)** | [Automated Deployment](#automated-deployment-recommended) | 5 min + 30 min Excel |
| **🔧 Manual setup (existing VM)** | [Manual Installation](#manual-installation) | 15 min + 30 min Excel |
| **📖 Infrastructure details** | [`infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md`](../infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md) | Reference |
| **🔍 Infrastructure code** | [`infrastructure/azure/README.md`](../infrastructure/azure/README.md) | Reference |

---

## Overview

ExcelMcp requires Microsoft Excel for integration testing. GitHub-hosted runners don't include Excel, so integration tests are currently skipped in CI/CD. This guide shows how to set up an Azure Windows VM with Excel as a GitHub Actions self-hosted runner.

## Architecture

```
┌─────────────────────────────────────────────────────────┐
│ GitHub Repository                                        │
│                                                          │
│  ┌──────────────────────────────────────────┐          │
│  │ .github/workflows/integration-tests.yml  │          │
│  │ runs-on: [self-hosted, windows, excel]   │          │
│  └────────────────┬─────────────────────────┘          │
└───────────────────┼──────────────────────────────────────┘
                    │
                    ▼
┌─────────────────────────────────────────────────────────┐
│ Azure Windows VM                                         │
│                                                          │
│  ┌──────────────────────────────────────────┐          │
│  │ GitHub Actions Runner Service            │          │
│  │ - Windows Server 2022                    │          │
│  │ - .NET 8 SDK                             │          │
│  │ - Microsoft Excel (Office 365)           │          │
│  │ - Self-hosted runner agent               │          │
│  └──────────────────────────────────────────┘          │
└─────────────────────────────────────────────────────────┘
```

---

## Automated Deployment (Recommended)

**✨ Fastest way to deploy - only manual step is installing Excel!**

**What gets automated:**
- ✅ VM provisioning (Standard_B2s, 4GB RAM - cheapest suitable option)
- ✅ .NET 8 SDK installation
- ✅ GitHub Actions runner installation & configuration
- ✅ Network security configuration
- ⏭️ **Manual:** Office 365 Excel installation (you must do this via RDP)

**Complete guide:** [`infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md`](../infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md)

**Cost:** ~$30/month (with auto-shutdown) or ~$60/month (24/7) in East US region

---

## Manual Installation

**Use this option if:**
- Automated deployment workflow failed
- You already have a Windows VM
- You want complete control over the setup

### Prerequisites

- Windows Server 2022 or Windows 10/11 VM (Azure or on-premises)
- Administrator access to the VM via RDP
- VM has internet connectivity
- Office 365 subscription with Excel license

### Installation Steps

#### 1. Connect to VM via RDP

Get your VM's public IP from Azure Portal, then:
```
Computer: <VM_PUBLIC_IP>
Username: Your admin username
Password: Your admin password
```

#### 2. Install .NET 8 SDK

Open PowerShell as Administrator:

```powershell
# Download .NET 8 SDK
Invoke-WebRequest -Uri "https://aka.ms/dotnet/8.0/dotnet-sdk-win-x64.exe" -OutFile "$env:TEMP\dotnet-sdk.exe"

# Install silently
Start-Process "$env:TEMP\dotnet-sdk.exe" -ArgumentList '/quiet' -Wait

# Verify
dotnet --version
```

#### 3. Generate GitHub Runner Token

**Important:** Tokens expire after 1 hour!

1. Go to repository: `https://github.com/sbroenne/mcp-server-excel`
2. Navigate to **Settings** → **Actions** → **Runners**
3. Click **New self-hosted runner**
4. Select **Windows**
5. Copy the registration token (long alphanumeric string)

#### 4. Download and Configure GitHub Actions Runner

In PowerShell as Administrator:

```powershell
# Create runner directory
New-Item -Path C:\actions-runner -ItemType Directory -Force
Set-Location C:\actions-runner

# Download latest runner
$runnerVersion = "2.321.0"  # Check GitHub for latest version
Invoke-WebRequest -Uri "https://github.com/actions/runner/releases/download/v$runnerVersion/actions-runner-win-x64-$runnerVersion.zip" -OutFile "actions-runner.zip"

# Extract
Expand-Archive -Path actions-runner.zip -DestinationPath . -Force

# Configure (replace with your token from Step 3)
$githubToken = "PASTE_YOUR_TOKEN_HERE"
$repoUrl = "https://github.com/sbroenne/mcp-server-excel"

.\config.cmd --url $repoUrl --token $githubToken --name "azure-excel-runner" --labels "self-hosted,windows,excel" --runnergroup Default --work _work --unattended
```

#### 5. Install Runner as Windows Service

```powershell
# Install service
.\svc.cmd install

# Start service
.\svc.cmd start

# Verify
Get-Service actions.runner.*
# Should show: Running
```

#### 6. Install Office 365 Excel

**Manual installation required:**

1. Open browser on VM → `https://portal.office.com`
2. Sign in with Office 365 account
3. Click **Install Office** → **Office 365 apps**
4. During installation, select **Excel only**
5. Complete installation (~15-30 minutes)
6. Open Excel once to activate (File → Account → verify activation)

#### 7. Verify Excel COM Access

```powershell
try {
    $excel = New-Object -ComObject Excel.Application
    $version = $excel.Version
    Write-Host "✅ Excel Version: $version" -ForegroundColor Green
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
} catch {
    Write-Host "❌ Excel not accessible: $_" -ForegroundColor Red
}
```

Expected: `✅ Excel Version: 16.0`

#### 8. Verify Runner Registration

Check `https://github.com/sbroenne/mcp-server-excel/settings/actions/runners`:
- **Name:** azure-excel-runner
- **Status:** Idle (green circle)
- **Labels:** self-hosted, windows, excel

#### 9. Test Integration Tests

1. Go to **Actions** tab → **Integration Tests (Excel)**
2. Click **Run workflow** → select `main` branch
3. Monitor the run - should complete successfully

### Manual Installation Troubleshooting

**Runner service won't start:**
```powershell
Get-EventLog -LogName Application -Source actions.runner.* -Newest 20
```

**"Runner already exists" error:**
```powershell
.\config.cmd remove --token YOUR_NEW_TOKEN
# Then reconfigure with Step 4 commands
```

**Excel COM test fails:**
- Verify Excel is installed and activated
- Kill background processes: `Get-Process excel | Stop-Process -Force`

**Runner token expired:**
- Generate new token (Step 3) and reconfigure

---

## Cost Estimate

## Cost Estimate

**Monthly costs (East US region - cheapest):**

| Resource | Specification | Monthly Cost (USD) |
|----------|---------------|-------------------|
| VM (Standard_B2s) | 2 vCPUs, 4 GB RAM | ~$25 |
| Storage (Premium SSD) | 128 GB | ~$5 |
| Network Egress | ~10 GB/month | <$1 |
| **Total (with auto-shutdown)** | | **~$30/month** |

**Other VM options:**
- Standard_B2ms (2 vCPUs, 8 GB): ~$60/month
- Standard_D2s_v3 (2 vCPUs, 8 GB): ~$70/month

**Cost optimization:**
- ✅ Use B2s (cheapest suitable VM)
- ✅ Enable auto-shutdown at 7 PM (saves ~50%)
- ✅ Use East US region (cheapest)
- Deallocate when not in use: ~$5/month (storage only)

---

## Azure Portal VM Creation (Optional)

If you prefer using Azure Portal instead of automation:

## Azure Portal VM Creation (Optional)

If you prefer using Azure Portal instead of automation:

1. Sign in to https://portal.azure.com
2. Create a resource → Virtual Machine
3. Configure:
   - Resource Group: `rg-excel-runner`
   - VM Name: `vm-excel-runner-01`
   - Region: East US (cheapest)
   - Image: Windows Server 2022 Datacenter
   - Size: Standard_B2s (2 vCPUs, 4 GB RAM)
   - Authentication: Set username/password
4. Networking: Allow RDP (3389)
5. Management: Enable auto-shutdown at 7 PM
6. Review + Create

Then follow [Manual Installation](#manual-installation) steps above.

---

## Maintenance & Operations

### Monitor Runner Health

**Check runner status:**
```powershell
# PowerShell on VM
Get-Service actions.runner.* | Format-Table Name, Status, StartType
```

**View runner logs:**
```powershell
# On VM
Get-Content "C:\actions-runner\_diag\Runner_*.log" -Tail 50
```

**GitHub Portal:**
- Go to: Settings → Actions → Runners
- Verify runner shows as "Idle" (green) or "Active" (running job)

### Update Runner

**When new runner version released:**
```powershell
# Stop service
.\svc.cmd stop

# Download new version
$newVersion = "2.322.0"  # Check GitHub for latest
Invoke-WebRequest -Uri "https://github.com/actions/runner/releases/download/v$newVersion/actions-runner-win-x64-$newVersion.zip" -OutFile "actions-runner-new.zip"

# Backup old version
Rename-Item "actions-runner.zip" "actions-runner-old.zip"
Rename-Item "actions-runner-new.zip" "actions-runner.zip"

# Extract (overwrites files)
[System.IO.Compression.ZipFile]::ExtractToDirectory("$PWD\actions-runner.zip", "$PWD")

# Restart service
.\svc.cmd start
```

### Cleanup Excel Processes

**After failed tests:**
```powershell
# Kill all Excel processes
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force

# Verify no orphan processes
Get-Process | Where-Object { $_.ProcessName -like "*excel*" -or $_.ProcessName -like "*dotnet*" }
```

### Auto-Shutdown Schedule

**Modify shutdown time:**
```powershell
# Azure Portal
# VM → Auto-shutdown → Change time → Save

# Azure CLI
az vm auto-shutdown --resource-group rg-excel-runner --name vm-excel-runner-01 --time 1900  # 7 PM
```

### Backup Runner Configuration

**Before major changes:**
```powershell
# Backup runner config
Copy-Item "C:\actions-runner\.runner" "C:\Backup\.runner.bak"
Copy-Item "C:\actions-runner\.credentials" "C:\Backup\.credentials.bak"
```

---

## Troubleshooting

### Runner Token Generation Fails

**Symptoms:** Automated deployment workflow fails with "Failed to generate runner registration token" or "Resource not accessible by integration" (403 error)

**Root Cause:** The `GITHUB_TOKEN` cannot create runner registration tokens via direct REST API calls, even with `actions: write` permission. This is a GitHub security restriction.

**Solution:** Use GitHub CLI instead of curl

**Before (Failed):**
```bash
curl -L -X POST \
  -H "Authorization: Bearer ${{ secrets.GITHUB_TOKEN }}" \
  https://api.github.com/repos/.../actions/runners/registration-token
```

**After (Fixed):**
```bash
gh api --method POST \
  /repos/${{ github.repository }}/actions/runners/registration-token \
  --jq '.token'
```

**Why It Works:** The GitHub CLI (`gh`) has proper authentication mechanisms that work with runner operations, while direct API calls are blocked for security reasons.

**Verification:** The automated deployment workflow (`.github/workflows/deploy-azure-runner.yml`) already uses this fix. If you're implementing manual deployment, use `gh api` instead of `curl` for token generation.

### Runner Not Appearing in GitHub

**Symptoms:** Runner not listed in Settings → Actions → Runners

**Solutions:**
1. Check service status: `Get-Service actions.runner.*`
2. Restart service: `.\svc.cmd restart`
3. View logs: `Get-Content "C:\actions-runner\_diag\Runner_*.log" -Tail 100`
4. Verify token expiration (tokens expire after 1 hour) - regenerate and reconfigure
5. Check network connectivity: `Test-NetConnection github.com -Port 443`

### Integration Tests Failing

**Symptoms:** Tests pass locally but fail on runner

**Solutions:**

1. **Excel not activated:**
   ```powershell
   # Launch Excel manually once
   Start-Process excel -Wait
   # Sign in with Office 365 account
   ```

2. **VBA trust not enabled:**
   ```powershell
   # Set VBA trust registry key
   Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" -Name "AccessVBOM" -Value 1
   ```

3. **Protected view blocking files:**
   ```powershell
   # Disable protected view
   $pvPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView"
   Set-ItemProperty -Path $pvPath -Name "DisableInternetFilesInPV" -Value 1
   ```

4. **Excel processes not cleaned up:**
   ```powershell
   # Add cleanup step to workflow
   Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
   ```

### RDP Connection Issues

**Cannot connect to VM:**

1. **Check VM is running:**
   ```bash
   az vm get-instance-view --resource-group rg-excel-runner --name vm-excel-runner-01 --query "instanceView.statuses[?starts_with(code, 'PowerState/')].displayStatus" -o tsv
   ```

2. **Start VM if stopped:**
   ```bash
   az vm start --resource-group rg-excel-runner --name vm-excel-runner-01
   ```

3. **Verify NSG rules allow your IP:**
   ```bash
   az network nsg rule list --resource-group rg-excel-runner --nsg-name vm-excel-runner-01NSG --query "[?name=='RDP'].{Name:name,Priority:priority,SourceAddressPrefix:sourceAddressPrefix}" -o table
   ```

4. **Update NSG rule to allow your current IP:**
   ```bash
   MY_IP=$(curl -s https://api.ipify.org)
   az network nsg rule update --resource-group rg-excel-runner --nsg-name vm-excel-runner-01NSG --name RDP --source-address-prefix "$MY_IP/32"
   ```

### High Azure Costs

**Monthly bill higher than expected:**

1. **Check VM running time:**
   - Azure Portal → Cost Management → Cost analysis
   - Filter by VM resource

2. **Verify auto-shutdown working:**
   ```bash
   az vm show --resource-group rg-excel-runner --name vm-excel-runner-01 --query "autoShutdownConfiguration"
   ```

3. **Stop VM completely when not needed:**
   ```bash
   az vm stop --resource-group rg-excel-runner --name vm-excel-runner-01
   az vm deallocate --resource-group rg-excel-runner --name vm-excel-runner-01  # Important: Deallocate to stop compute billing
   ```

4. **Downgrade VM size if underutilized:**
   ```bash
   # Resize to B2s (cheapest)
   az vm resize --resource-group rg-excel-runner --name vm-excel-runner-01 --size Standard_B2s
   ```

### Excel Automation Errors

**Tests failing with COM errors:**

1. **DCOM permissions:**
   ```powershell
   # Run as Administrator
   dcomcnfg
   # Component Services → Computers → My Computer → DCOM Config → Microsoft Excel Application
   # Right-click → Properties → Identity → The interactive user
   ```

2. **Excel hanging:**
   ```powershell
   # Add timeout to test configuration
   # In test code: Disable background save, disable add-ins
   ```

3. **File locks:**
   ```powershell
   # Ensure tests dispose Excel objects properly
   # Check for orphan Excel processes: Get-Process excel
   ```

---

## Cleanup & Decommissioning

### Remove Runner from GitHub

**Before deleting VM:**
```powershell
# On VM - stop and remove service
.\svc.cmd stop
.\svc.cmd uninstall

# Unregister from GitHub
.\config.cmd remove --token YOUR_REMOVAL_TOKEN
```

**GitHub Portal:**
- Settings → Actions → Runners
- Click runner name → **Remove runner**

### Delete Azure Resources

**Remove all infrastructure:**
```bash
# Delete resource group (removes VM, disk, network, etc.)
az group delete --name rg-excel-runner --yes --no-wait
```

**Verify deletion:**
```bash
az group list --query "[?name=='rg-excel-runner']" -o table
```

### Cost After Deletion

After deletion: **$0/month** (all resources removed)

If you only stop VM: **~$5/month** (storage costs remain)

---

## Security Best Practices

### Minimize Attack Surface

1. **Restrict RDP access to your IP only** (see [Configure Network Security](#step-4-configure-network-security))
2. **Use strong admin password** (16+ characters, mixed case, numbers, symbols)
3. **Enable auto-shutdown** to reduce exposure time
4. **Keep Windows updated:**
   ```powershell
   # Check for updates
   Install-Module PSWindowsUpdate -Force
   Get-WindowsUpdate
   Install-WindowsUpdate -AcceptAll -AutoReboot
   ```

### GitHub Secrets Management

- **Never commit runner tokens** to repository
- Use GitHub Secrets for sensitive workflow inputs
- Rotate registration tokens regularly (they expire after 1 hour anyway)

### Monitor for Suspicious Activity

**Azure Security Center:**
- Enable Microsoft Defender for Cloud (free tier)
- Review security recommendations
- Monitor alerts for brute-force attempts on RDP

**GitHub:**
- Review workflow runs for unexpected triggers
- Monitor runner logs for unauthorized jobs

### Principle of Least Privilege

- Runner service account should only have permissions needed for tests
- Don't run runner as domain admin or with elevated privileges
- Restrict file system access to runner work directory

---

## Cost Optimization Strategies

### Strategy 1: Auto-Shutdown (Recommended)

**Setup:**
- Enable auto-shutdown at 7 PM (or your EOD)
- Manually start VM before running scheduled tests
- Saves ~50% compute costs

**Best for:** Teams in single time zone, predictable schedules

### Strategy 2: Start/Stop VM with Automation

**PowerShell script (local machine):**
```powershell
# start-runner.ps1
az vm start --resource-group rg-excel-runner --name vm-excel-runner-01

# Wait for VM to start
Start-Sleep -Seconds 60

# Trigger integration tests (via workflow_dispatch API or push)
# ...

# stop-runner.ps1 (after tests complete)
az vm stop --resource-group rg-excel-runner --name vm-excel-runner-01
az vm deallocate --resource-group rg-excel-runner --name vm-excel-runner-01
```

**Best for:** Infrequent test runs, CI/CD pipelines

### Strategy 3: Use Spot Instances (Advanced)

**Lower cost but VM can be evicted:**
```bash
az vm create \
  --resource-group rg-excel-runner \
  --name vm-excel-runner-spot \
  --priority Spot \
  --max-price 0.05 \
  --eviction-policy Deallocate \
  # ... other parameters
```

**Best for:** Non-critical test runs, can tolerate interruptions

### Strategy 4: Resize Based on Load

**Scale up for heavy workloads:**
```bash
# Before intensive tests
az vm resize --resource-group rg-excel-runner --name vm-excel-runner-01 --size Standard_D2s_v3

# After tests complete
az vm resize --resource-group rg-excel-runner --name vm-excel-runner-01 --size Standard_B2s
```

**Best for:** Occasional heavy workloads, cost-sensitive projects

---

## References

- [GitHub Actions Self-Hosted Runners](https://docs.github.com/en/actions/hosting-your-own-runners/about-self-hosted-runners)
- [Azure Windows VMs Pricing](https://azure.microsoft.com/en-us/pricing/details/virtual-machines/windows/)
- [Excel COM Automation](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
- [Azure Auto-Shutdown](https://docs.microsoft.com/en-us/azure/virtual-machines/auto-shutdown-vm)

---

## Quick Reference Commands

**Start/Stop VM:**
```bash
az vm start --resource-group rg-excel-runner --name vm-excel-runner-01
az vm stop --resource-group rg-excel-runner --name vm-excel-runner-01
az vm deallocate --resource-group rg-excel-runner --name vm-excel-runner-01
```

**Check runner status:**
```powershell
Get-Service actions.runner.*
```

**Restart runner service:**
```powershell
.\svc.cmd restart
```

**View runner logs:**
```powershell
Get-Content "C:\actions-runner\_diag\Runner_*.log" -Tail 50
```

**Kill Excel processes:**
```powershell
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
```

**Check Azure costs:**
```bash
az consumption usage list --resource-group rg-excel-runner --query "[].{Resource:instanceName,Cost:pretaxCost}" -o table
```

### Start/Stop VM

**Azure Portal:**
- Navigate to VM → Click **Start** or **Stop**

**Azure CLI:**
```bash
# Stop VM (deallocate to save costs)
az vm deallocate --resource-group rg-excel-runner --name vm-excel-runner-01

# Start VM
az vm start --resource-group rg-excel-runner --name vm-excel-runner-01
```

### Auto-Shutdown Schedule

**Azure Portal:**
1. Go to VM → **Auto-shutdown**
2. Enable: **On**
3. Shutdown time: `19:00` (7 PM)
4. Time zone: Your local timezone
5. Notification: Configure email (optional)
6. **Save**

### Update Runner

**PowerShell (on VM):**
```powershell
# Stop runner service
C:\actions-runner\svc.cmd stop

# Download latest version
$runnerVersion = "2.321.0"  # Update to latest
Invoke-WebRequest -Uri "https://github.com/actions/runner/releases/download/v$runnerVersion/actions-runner-win-x64-$runnerVersion.zip" -OutFile "C:\actions-runner\actions-runner-new.zip"

# Extract to temp location
Expand-Archive -Path "C:\actions-runner\actions-runner-new.zip" -DestinationPath "C:\actions-runner-new" -Force

# Replace binaries (preserve config)
Copy-Item -Path "C:\actions-runner-new\*" -Destination "C:\actions-runner\" -Recurse -Force -Exclude ".credentials",".runner"

# Start runner service
C:\actions-runner\svc.cmd start
```

### Monitor Runner Health

**Check runner status:**
```powershell
# On VM
Get-Service actions.runner.* | Select-Object Name, Status, StartType

# View logs
Get-Content "C:\actions-runner\_diag\Runner*.log" -Tail 50
```

**GitHub UI:**
- Go to repository **Settings** → **Actions** → **Runners**
- Check runner status (Idle/Active/Offline)

## Troubleshooting

### Runner Shows Offline

**Check service status:**
```powershell
Get-Service actions.runner.*
# If stopped, restart:
Restart-Service actions.runner.*
```

**Check network connectivity:**
```powershell
Test-NetConnection -ComputerName github.com -Port 443
```

### Excel COM Errors in Tests

**Verify Excel is installed:**
```powershell
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Excel*" }
```

**Check Excel process cleanup:**
```powershell
# Kill orphaned Excel processes
Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
```

### Tests Timeout

- Increase `timeout-minutes` in workflow
- Check VM performance (CPU/RAM usage)
- Consider upgrading VM size

### Licensing Issues

- Ensure Office 365 license is active
- Re-activate Excel if needed:
  ```powershell
  Start-Process excel
  # Sign in interactively via RDP
  ```

## Security Best Practices

1. **Restrict Runner to Private Repos Only**
   - Go to **Settings** → **Actions** → **Runner groups**
   - Ensure runner group only allows private repositories

2. **Use Dedicated Service Account**
   - Create Azure AD user specifically for runner
   - Apply principle of least privilege

3. **Regular Updates**
   - Enable Windows Update
   - Update runner agent monthly
   - Update Excel/Office monthly

4. **Secrets Management**
   - Never hardcode credentials in workflows
   - Use GitHub Secrets for sensitive data
   - Rotate runner registration tokens

5. **Network Isolation**
   - Use Azure Bastion instead of RDP (enterprise)
   - Restrict NSG to minimum required ports
   - Consider private VNet for runner

## Alternative Solutions

### Option 1: Azure Container Apps (Future)

Microsoft is developing container-based CI/CD runners that could potentially support Windows containers with Excel. Monitor [this announcement](https://learn.microsoft.com/en-us/azure/container-apps/tutorial-ci-cd-runners-jobs).

### Option 2: Azure Virtual Desktop Multi-Session

For multiple concurrent test runs, consider Azure Virtual Desktop with multi-session host pools.

### Option 3: Third-Party Hosted Runners

Some CI/CD providers offer Windows runners with Office pre-installed:
- **BuildJet** (GitHub Actions accelerator with custom images)
- **Cirrus CI** (Windows containers with Office)

Cost comparison needed before adoption.

## Cost Optimization Strategies

1. **Scheduled Start/Stop**
   - Use Azure Automation runbooks
   - Start VM 30 min before scheduled test run
   - Stop VM after tests complete

2. **Spot VMs**
   - Save up to 90% on VM costs
   - Acceptable for non-critical test runs
   - Risk: VM can be evicted by Azure

3. **Reserved Instances**
   - 1-year commitment: ~40% savings
   - 3-year commitment: ~60% savings
   - Only if runner runs 24/7

4. **B-Series Burstable VMs**
   - Lower base cost
   - Suitable for intermittent workloads
   - May impact test performance

## Next Steps

After setup:

1. ✅ Test runner with simple workflow
2. ✅ Run integration tests manually
3. ✅ Configure auto-shutdown to reduce costs
4. ✅ Set up monitoring/alerting
5. ✅ Document runner in team wiki

## Additional Resources

- [GitHub Self-Hosted Runners Documentation](https://docs.github.com/en/actions/hosting-your-own-runners/managing-self-hosted-runners/about-self-hosted-runners)
- [Azure Virtual Machines Documentation](https://learn.microsoft.com/en-us/azure/virtual-machines/)
- [Office Deployment Tool](https://learn.microsoft.com/en-us/deployoffice/overview-office-deployment-tool)
- [Azure Cost Management](https://azure.microsoft.com/en-us/products/cost-management/)

## Support

For issues or questions:
- GitHub Issues: https://github.com/sbroenne/mcp-server-excel/issues
- Documentation: [DEVELOPMENT.md](DEVELOPMENT.md)
