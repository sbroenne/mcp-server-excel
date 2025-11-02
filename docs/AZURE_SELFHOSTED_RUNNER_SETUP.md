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

**Monthly costs (East US region - cheapest):**

| Resource | Specification | Monthly Cost (USD) |
|----------|---------------|-------------------|
| VM (Standard_B2s) | 2 vCPUs, 4 GB RAM | ~$25 |
| Storage (Premium SSD) | 128 GB | ~$5 |
| Network Egress | ~10 GB/month | <$1 |
| **Total (with auto-shutdown)** | | **~$30/month** |

**Other VM Options:**
- Standard_B2ms (2 vCPUs, 8 GB): ~$60/month
- Standard_D2s_v3 (2 vCPUs, 8 GB): ~$70/month

**Cost Optimization:**
- ✅ **Automated deployment uses B2s** (cheapest suitable VM)
- ✅ **East US region** (cheapest region)
- ✅ **Auto-shutdown at 7 PM** (saves ~50%)
- Stop VM completely: ~$5/month (storage only)

> 💡 **Recommended:** Use automated deployment with B2s + auto-shutdown = $30/month

## Step 1: Create Azure Windows VM

### Option A: Azure Portal (Manual)

1. **Sign in to Azure Portal**: https://portal.azure.com

2. **Create Virtual Machine:**
   - Click **Create a resource** → **Virtual Machine**
   - **Basics:**
     - Subscription: Select your subscription
     - Resource Group: Create new `rg-excel-runner`
     - VM Name: `vm-excel-runner-01`
     - Region: Choose closest to your location
     - Image: **Windows Server 2022 Datacenter**
     - Size: **Standard_D2s_v3** (2 vCPUs, 8 GB RAM)
     - Username: `adminuser` (choose your own)
     - Password: Strong password (save securely)
   
   - **Disks:**
     - OS Disk Type: **Premium SSD** (128 GB)
   
   - **Networking:**
     - Virtual Network: Create new or use existing
     - Public IP: **Create new** (for RDP access)
     - NIC Security Group: **Basic**
     - Public Inbound Ports: **RDP (3389)**
   
   - **Management:**
     - Enable Auto-shutdown: **Yes** (e.g., 7:00 PM daily)
   
   - Click **Review + Create** → **Create**

3. **Wait for deployment** (~5 minutes)

### Option B: Azure CLI (Automated)

```bash
# Login to Azure
az login

# Create resource group
az group create --name rg-excel-runner --location eastus

# Create VM
az vm create \
  --resource-group rg-excel-runner \
  --name vm-excel-runner-01 \
  --image Win2022Datacenter \
  --size Standard_D2s_v3 \
  --admin-username adminuser \
  --admin-password 'YourSecurePassword123!' \
  --public-ip-sku Standard \
  --os-disk-size-gb 128

# Open RDP port
az vm open-port --port 3389 --resource-group rg-excel-runner --name vm-excel-runner-01

# Get public IP
az vm show -d --resource-group rg-excel-runner --name vm-excel-runner-01 --query publicIps -o tsv
```

## Step 2: Connect to VM and Install Prerequisites

### Connect via RDP

1. Get VM Public IP from Azure Portal
2. Open **Remote Desktop Connection** on your local machine
3. Enter `<Public-IP>` and credentials
4. Click **Connect**

### Install .NET 8 SDK

**PowerShell (as Administrator):**
```powershell
# Download .NET 8 SDK
Invoke-WebRequest -Uri https://aka.ms/dotnet/8.0/dotnet-sdk-win-x64.exe -OutFile dotnet-sdk-8.0.exe

# Install silently
Start-Process -FilePath .\dotnet-sdk-8.0.exe -ArgumentList '/quiet', '/norestart' -Wait

# Verify installation
dotnet --version
```

### Install Microsoft Excel (Office 365)

**Option 1: Office Deployment Tool (Recommended)**

```powershell
# Download Office Deployment Tool
Invoke-WebRequest -Uri https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117 -OutFile officedeploymenttool.exe

# Extract
Start-Process -FilePath .\officedeploymenttool.exe -ArgumentList '/extract:ODT' -Wait

# Create configuration.xml
@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="PowerPoint" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="Word" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@ | Out-File -FilePath .\ODT\configuration.xml

# Install Excel only
.\ODT\setup.exe /configure .\ODT\configuration.xml
```

**Option 2: Manual Installation**

1. Sign in with Office 365 account at https://portal.office.com
2. Click **Install Office** → **Office 365 apps**
3. Run installer on the VM
4. During installation, choose **Customize** and select **Excel only**

### Activate Excel

```powershell
# Launch Excel once to activate
Start-Process excel -Wait

# Accept license agreement and sign in with Office 365 account
# (This step must be done interactively via RDP)
```

### Configure Excel for Automation

**PowerShell (as Administrator):**
```powershell
# Disable Excel splash screen and startup tasks
$excelPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
New-Item -Path $excelPath -Force | Out-Null
Set-ItemProperty -Path $excelPath -Name "DisableBootToOfficeStart" -Value 1
Set-ItemProperty -Path $excelPath -Name "BootedRTM" -Value 1

# Trust VBA project access (required for VBA tests)
$trustPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
New-Item -Path $trustPath -Force | Out-Null
Set-ItemProperty -Path $trustPath -Name "AccessVBOM" -Value 1
Set-ItemProperty -Path $trustPath -Name "VBAWarnings" -Value 1

# Disable protected view (for test files)
$pvPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\ProtectedView"
New-Item -Path $pvPath -Force | Out-Null
Set-ItemProperty -Path $pvPath -Name "DisableInternetFilesInPV" -Value 1
Set-ItemProperty -Path $pvPath -Name "DisableAttachmentsInPV" -Value 1
Set-ItemProperty -Path $pvPath -Name "DisableUnsafeLocationsInPV" -Value 1
```

## Step 3: Install GitHub Actions Runner

### Generate Registration Token

1. Go to GitHub repository: `https://github.com/sbroenne/mcp-server-excel`
2. Navigate to **Settings** → **Actions** → **Runners**
3. Click **New self-hosted runner**
4. Select **Windows** platform
5. Copy the registration token (valid for 1 hour)

### Install Runner on VM

**PowerShell (as Administrator):**

```powershell
# Create runner directory
New-Item -Path "C:\actions-runner" -ItemType Directory
Set-Location "C:\actions-runner"

# Download latest runner (check GitHub for latest version)
$runnerVersion = "2.321.0"  # Update to latest version
Invoke-WebRequest -Uri "https://github.com/actions/runner/releases/download/v$runnerVersion/actions-runner-win-x64-$runnerVersion.zip" -OutFile "actions-runner.zip"

# Extract
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory("$PWD\actions-runner.zip", "$PWD")

# Configure runner
$GITHUB_REPO_URL = "https://github.com/sbroenne/mcp-server-excel"
$REGISTRATION_TOKEN = "YOUR_TOKEN_HERE"  # Replace with token from GitHub

.\config.cmd --url $GITHUB_REPO_URL --token $REGISTRATION_TOKEN --name "azure-excel-runner" --labels "self-hosted,windows,excel" --runnergroup "Default" --work "_work" --unattended

# Install and start as Windows service (runs on boot)
.\svc.cmd install
.\svc.cmd start

# Verify service is running
Get-Service actions.runner.*
```

### Test Runner Connection

1. Go back to GitHub: **Settings** → **Actions** → **Runners**
2. You should see `azure-excel-runner` with status **Idle** and labels: `self-hosted`, `windows`, `excel`

## Step 4: Configure Network Security

### Restrict RDP Access (Recommended)

**Azure Portal:**
1. Go to VM → **Networking** → **Network settings**
2. Find RDP rule (port 3389)
3. Click **Edit** → **Source**: Change from `Any` to `My IP address`
4. **Save**

**Azure CLI:**
```bash
# Get your public IP
MY_IP=$(curl -s https://api.ipify.org)

# Update NSG rule
az vm open-port --port 3389 --resource-group rg-excel-runner --name vm-excel-runner-01 --priority 1000 --source-address-prefix "$MY_IP/32"
```

### Firewall Rules (Optional)

The runner uses HTTPS (443) for GitHub communication - already allowed by default Azure NSG.

## Step 5: Create Integration Test Workflow

Create new file: `.github/workflows/integration-tests.yml`

```yaml
name: Integration Tests (Excel)

on:
  # Run on schedule (e.g., nightly)
  schedule:
    - cron: '0 2 * * *'  # 2 AM UTC daily
  
  # Allow manual trigger
  workflow_dispatch:
  
  # Optionally run on PR to main (only if you want to block merges)
  # pull_request:
  #   branches: [ main ]

jobs:
  integration-tests:
    runs-on: [self-hosted, windows, excel]
    timeout-minutes: 60
    
    steps:
    - name: Checkout code
      uses: actions/checkout@v4
    
    - name: Setup .NET
      uses: actions/setup-dotnet@v4
      with:
        dotnet-version: 8.0.x
    
    - name: Restore dependencies
      run: dotnet restore
    
    - name: Build
      run: dotnet build --no-restore --configuration Release
    
    - name: Run Integration Tests
      run: dotnet test --no-build --configuration Release --filter "Category=Integration&RunType!=OnDemand" --logger "trx;LogFileName=integration-test-results.trx"
    
    - name: Upload Test Results
      if: always()
      uses: actions/upload-artifact@v4
      with:
        name: integration-test-results
        path: '**/TestResults/*.trx'
    
    - name: Cleanup Excel Processes
      if: always()
      run: |
        Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 5
      shell: pwsh
```

## Step 6: Update Existing Workflows (Optional)

Add integration test status badge to `README.md`:

```markdown
[![Integration Tests](https://github.com/sbroenne/mcp-server-excel/actions/workflows/integration-tests.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/integration-tests.yml)
```

Reference integration tests from main build workflows:

```yaml
# In build-mcp-server.yml or build-cli.yml
jobs:
  build:
    # ... existing build job ...
  
  trigger-integration-tests:
    needs: build
    if: github.event_name == 'push' && github.ref == 'refs/heads/main'
    runs-on: ubuntu-latest
    steps:
      - name: Trigger Integration Tests
        uses: peter-evans/repository-dispatch@v3
        with:
          event-type: trigger-integration-tests
          token: ${{ secrets.GITHUB_TOKEN }}
```

## Maintenance & Operations

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
