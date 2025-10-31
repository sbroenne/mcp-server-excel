# Manual GitHub Runner Installation for Excel Testing

> **When to use this guide:** If the automated Azure deployment workflow failed or you have an existing Windows VM that you want to configure as a GitHub runner

## Prerequisites

- Windows Server 2022 or Windows 10/11 VM (already provisioned in Azure)
- Administrator access to the VM via RDP
- VM has internet connectivity
- Office 365 subscription (for Excel)

## Step-by-Step Installation

### Step 1: RDP to Your VM

1. Get your VM's public IP or FQDN from Azure Portal
2. Open Remote Desktop Connection
3. Connect using:
   - **Computer:** `<VM_PUBLIC_IP_OR_FQDN>`
   - **Username:** `azureuser` (or your configured admin username)
   - **Password:** `<your_admin_password>`

### Step 2: Install .NET 8 SDK

Open PowerShell as Administrator and run:

```powershell
# Download .NET 8 SDK installer
Invoke-WebRequest -Uri "https://aka.ms/dotnet/8.0/dotnet-sdk-win-x64.exe" -OutFile "$env:TEMP\dotnet-sdk.exe"

# Install silently
Start-Process "$env:TEMP\dotnet-sdk.exe" -ArgumentList '/quiet' -Wait

# Verify installation
dotnet --version
# Should output: 8.0.x
```

### Step 3: Generate GitHub Runner Token

**Important:** Runner tokens expire after 1 hour, so generate this right before Step 4!

1. Go to your repository on GitHub: `https://github.com/sbroenne/mcp-server-excel`
2. Navigate to **Settings** → **Actions** → **Runners**
3. Click **New self-hosted runner**
4. Select **Windows** as the operating system
5. Copy the registration token from the configuration command
   - Token format: Long alphanumeric string starting with 'A'
   - Example: `A3E7G2K...` (keep this secret!)

### Step 4: Download and Configure GitHub Actions Runner

In PowerShell as Administrator:

```powershell
# Create runner directory
New-Item -Path C:\actions-runner -ItemType Directory -Force
Set-Location C:\actions-runner

# Download the latest runner package
$runnerVersion = "2.321.0"
Invoke-WebRequest -Uri "https://github.com/actions/runner/releases/download/v$runnerVersion/actions-runner-win-x64-$runnerVersion.zip" -OutFile "actions-runner.zip"

# Extract the installer
Expand-Archive -Path actions-runner.zip -DestinationPath . -Force

# Configure the runner (REPLACE WITH YOUR TOKEN FROM STEP 3)
$githubToken = "PASTE_YOUR_TOKEN_HERE"
$repoUrl = "https://github.com/sbroenne/mcp-server-excel"

.\config.cmd --url $repoUrl --token $githubToken --name "azure-excel-runner" --labels "self-hosted,windows,excel" --runnergroup Default --work _work --unattended

# Expected output:
# ✓ Runner successfully added
# ✓ Runner connection is good
```

**Note:** If the token expired, go back to Step 3 to generate a new one.

### Step 5: Install Runner as Windows Service

In the same PowerShell window:

```powershell
# Install as service
.\svc.cmd install

# Start the service
.\svc.cmd start

# Verify service status
Get-Service actions.runner.*
# Should show: Running
```

### Step 6: Install Office 365 Excel

**Manual installation required** (cannot be automated):

1. Open browser on the VM
2. Go to `https://portal.office.com`
3. Sign in with your Office 365 account
4. Click **Install Office** → **Office 365 apps**
5. Run the installer
6. During installation, select **Excel only** (faster, saves disk space)
7. Complete the installation (takes ~15-30 minutes)
8. Open Excel once to activate:
   - File → Account
   - Verify activation status shows "Product Activated"

### Step 7: Verify Excel COM Access

Test that Excel is accessible via COM automation:

```powershell
# Test Excel COM
try {
    $excel = New-Object -ComObject Excel.Application
    $version = $excel.Version
    Write-Host "✅ Excel Version: $version" -ForegroundColor Green
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
} catch {
    Write-Host "❌ Excel not accessible: $_" -ForegroundColor Red
    exit 1
}
```

Expected output:
```
✅ Excel Version: 16.0
```

### Step 8: Verify Runner Registration

1. Go to `https://github.com/sbroenne/mcp-server-excel/settings/actions/runners`
2. You should see:
   - **Name:** azure-excel-runner
   - **Status:** Idle (green circle)
   - **Labels:** self-hosted, windows, excel
   - **Version:** 2.321.0

### Step 9: Test Integration Tests Workflow

Trigger the integration tests workflow:

1. Go to **Actions** tab in GitHub
2. Select **Integration Tests (Excel)** workflow
3. Click **Run workflow**
4. Select branch: `main`
5. Click **Run workflow**

Monitor the run - it should:
- ✅ Pick up the self-hosted runner
- ✅ Display Excel version
- ✅ Run integration tests
- ✅ Complete successfully

## Troubleshooting

### Runner Service Won't Start

```powershell
# Check service logs
Get-EventLog -LogName Application -Source actions.runner.* -Newest 20

# Try manual start with verbose logging
.\run.cmd
# Look for error messages, then press Ctrl+C and reinstall service
```

### "Runner already exists" Error

```powershell
# Remove existing runner configuration
.\config.cmd remove --token YOUR_NEW_TOKEN

# Reconfigure with new token (from Step 3)
.\config.cmd --url https://github.com/sbroenne/mcp-server-excel --token YOUR_NEW_TOKEN --name "azure-excel-runner" --labels "self-hosted,windows,excel" --unattended
```

### Excel COM Test Fails

**Possible causes:**
1. Excel not installed → Install from portal.office.com
2. Excel not activated → Open Excel, File → Account, sign in
3. Excel running in background → Kill processes: `Get-Process excel | Stop-Process -Force`

### Workflow Not Using Self-Hosted Runner

Check workflow file `.github/workflows/integration-tests.yml`:

```yaml
jobs:
  integration-tests:
    runs-on: [self-hosted, windows, excel]  # Must match runner labels
```

### Runner Token Expired

Generate a new token (Step 3) and reconfigure:

```powershell
cd C:\actions-runner
.\config.cmd remove --token NEW_TOKEN_HERE
.\config.cmd --url https://github.com/sbroenne/mcp-server-excel --token NEW_TOKEN_HERE --name "azure-excel-runner" --labels "self-hosted,windows,excel" --unattended
.\svc.cmd start
```

## Maintenance

### Update Runner

```powershell
cd C:\actions-runner
.\svc.cmd stop
# Download latest version from https://github.com/actions/runner/releases
# Extract to C:\actions-runner (replace files)
.\svc.cmd start
```

### Update Office/Excel

```powershell
# Open Excel → File → Account → Update Options → Update Now
# Or wait for automatic updates
```

### Check Runner Logs

```powershell
# View service logs
Get-Content C:\actions-runner\_diag\Runner_*.log -Tail 50

# View worker logs
Get-Content C:\actions-runner\_diag\Worker_*.log -Tail 50
```

### Restart Runner Service

```powershell
Restart-Service actions.runner.*
```

## Cleanup

To remove the runner:

```powershell
cd C:\actions-runner
.\svc.cmd stop
.\svc.cmd uninstall
.\config.cmd remove --token YOUR_TOKEN
```

Then delete the VM from Azure Portal if no longer needed.

## Security Best Practices

1. **Restrict RDP access** - Update Azure NSG to allow only your IP
2. **Use Azure Bastion** - For RDP without public IP exposure
3. **Rotate VM password** - Every 90 days
4. **Enable Windows Defender** - Keep real-time protection on
5. **Enable Windows Updates** - Auto-install security updates
6. **Use least privilege** - Runner runs as Network Service (secure)

## Cost Optimization

**Monthly cost estimate (Standard_B2ms, Sweden Central):**
- VM: ~$50/month
- Storage: ~$11/month
- Network: <$1/month
- **Total: ~$61/month (24/7)**

**Cost-saving options:**
- Use auto-shutdown (not recommended for CI/CD - causes delays)
- Use smaller VM size (B2s with 4GB RAM - may be unstable)
- Deallocate when not in use (manual, not recommended)

## Support

- **GitHub Runner Issues:** [GitHub Docs](https://docs.github.com/en/actions/hosting-your-own-runners)
- **Azure VM Issues:** [Azure Support](https://portal.azure.com/#blade/Microsoft_Azure_Support/HelpAndSupportBlade)
- **Excel/Office Issues:** [Office Support](https://support.microsoft.com/office)
- **Repository Issues:** [Create Issue](https://github.com/sbroenne/mcp-server-excel/issues/new)

## Next Steps

After successful setup:

1. ✅ Runner appears in GitHub Settings → Actions → Runners
2. ✅ Integration tests workflow runs successfully
3. ✅ Excel COM tests pass
4. ✅ Monitor costs in Azure Cost Management
5. ✅ Set up budget alerts ($40/month threshold recommended)

**Setup complete!** Your repository now has full Excel COM integration testing in CI/CD.
