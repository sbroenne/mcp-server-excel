# Azure Infrastructure for Excel Integration Testing

This directory contains Infrastructure as Code (IaC) for automating the deployment of Azure Windows VM with GitHub Actions self-hosted runner.

## Quick Start

### Prerequisites

- Azure CLI installed (`az --version`)
- Azure subscription with VM creation permissions
- GitHub repository admin access
- Office 365 E3/E5 license or standalone Excel

### Deploy in 5 Minutes

```bash
cd infrastructure/azure

# Login to Azure
az login

# Generate GitHub runner token
# Go to: https://github.com/sbroenne/mcp-server-excel/settings/actions/runners/new
# Select: Windows
# Copy the token from the configuration command

# Deploy (replace with your values)
./deploy.sh \
  rg-excel-runner \
  "YourSecureVMPassword123!" \
  "YOUR_GITHUB_RUNNER_TOKEN_HERE"

# RDP to VM and install Office 365 Excel (30 minutes)
# Runner auto-starts after reboot
```

## What Gets Deployed

### Resources Created

| Resource | Type | Purpose | Monthly Cost |
|----------|------|---------|--------------|
| VM | Standard_B2s (2 vCPUs, 4 GB RAM) | Test runner | ~$25 |
| OS Disk | Premium SSD 128 GB | Storage | ~$5 |
| Network | VNet, NIC, NSG, Public IP | Connectivity | <$1 |
| **Total** | | | **~$30/month** |

### Location & VM Size

- **Location:** East US (cheapest region)
- **VM Size:** Standard_B2s (cheapest burstable VM)
- **Auto-shutdown:** 7 PM UTC daily (saves ~50%)

### Software Installed Automatically

1. **.NET 8 SDK** - For building and testing
2. **GitHub Actions Runner** - Registered with `self-hosted`, `windows`, `excel` labels
3. **Auto-start service** - Runner starts on VM boot

### Manual Installation Required

**Office 365 Excel** (you must install this via RDP):
1. RDP to VM using public IP from deployment output
2. Sign in to https://portal.office.com
3. Install Office 365 apps
4. During installation, select **Excel only**
5. Activate with your Office 365 account
6. Reboot VM

## Files

```
infrastructure/azure/
├── azure-runner.bicep              # Main Bicep template
├── azure-runner.parameters.json    # Parameters (optional)
├── deploy.sh                       # Deployment script
└── README.md                       # This file
```

## Deployment Options

### Option 1: Bash Script (Recommended)

```bash
./deploy.sh rg-excel-runner "Password123!" "GITHUB_TOKEN"
```

### Option 2: Azure CLI Direct

```bash
az group create --name rg-excel-runner --location eastus

az deployment group create \
  --resource-group rg-excel-runner \
  --template-file azure-runner.bicep \
  --parameters \
    adminPassword="YourPassword123!" \
    githubRepoUrl="https://github.com/sbroenne/mcp-server-excel" \
    githubRunnerToken="YOUR_TOKEN"
```

### Option 3: Azure Portal

1. Upload `azure-runner.bicep` to Azure Portal
2. Fill in parameters
3. Review + Create
4. Deploy

## Configuration

### VM Size Options

| Size | vCPUs | RAM | Monthly Cost | Use Case |
|------|-------|-----|--------------|----------|
| **Standard_B2s** | 2 | 4 GB | ~$30 | Budget (recommended) ⭐ |
| Standard_B2ms | 2 | 8 GB | ~$60 | More memory |
| Standard_D2s_v3 | 2 | 8 GB | ~$70 | Better performance |

Edit `vmSize` parameter in `azure-runner.bicep` to change.

### Region Options (Cheapest to Most Expensive)

| Region | Cost Factor |
|--------|-------------|
| **East US** | 1.0x (cheapest) ⭐ |
| South Central US | 1.0x |
| West US 2 | 1.1x |
| North Europe | 1.2x |
| West Europe | 1.2x |

Edit `location` parameter to change.

### Auto-Shutdown Schedule

Default: 7 PM UTC (19:00)

Edit in `azure-runner.bicep`:
```bicep
dailyRecurrence: {
  time: '1900' // Change to '2000' for 8 PM, etc.
}
```

## Cost Optimization

### Current Setup (~$30/month)

- VM runs 12 hours/day (auto-shutdown at 7 PM)
- Total: ~$30/month

### Further Optimization

**Scheduled Start/Stop (2 hours/day): ~$5/month**

Use Azure Automation to:
- Start VM at 1:30 AM UTC
- Run tests at 2:00 AM UTC
- Stop VM at 3:00 AM UTC

**Weekend Shutdown:**
- Deallocate VM Friday night
- Start Monday morning
- Saves ~$8/month

## Verify Deployment

### Check Runner Status

```bash
# In GitHub UI
https://github.com/sbroenne/mcp-server-excel/settings/actions/runners

# Should show:
# - Name: azure-excel-runner
# - Status: Idle (green)
# - Labels: self-hosted, windows, excel
```

### Trigger Test Workflow

```bash
# In GitHub UI
Actions → Integration Tests (Excel) → Run workflow
```

### Monitor Costs

```bash
# Azure Cost Management
https://portal.azure.com/#view/Microsoft_Azure_CostManagement/Menu/~/overview

# Set budget alert at $40/month
```

## Maintenance

### Update Runner

SSH/RDP to VM:
```powershell
cd C:\actions-runner
.\svc.cmd stop
# Download latest runner from GitHub
# Extract and replace files
.\svc.cmd start
```

### Update Windows

Auto-updates enabled. Manual check:
```powershell
sconfig  # Option 6: Download and Install Updates
```

### Update Office/Excel

Office auto-updates enabled. Manual check:
```powershell
# Open Excel → File → Account → Update Options → Update Now
```

## Troubleshooting

### Runner Offline

```powershell
# On VM
Get-Service actions.runner.*
# If stopped:
Restart-Service actions.runner.*
```

### High Costs

1. Verify auto-shutdown is enabled
2. Check VM is deallocated when not in use
3. Review Azure Cost Management for unexpected resources

### Excel Activation Issues

1. RDP to VM
2. Open Excel
3. Sign in with Office 365 account
4. Verify activation: File → Account

## Security

### Network Access

- RDP (3389) restricted by NSG (change to your IP after deployment)
- HTTPS (443) outbound for GitHub
- All other ports blocked

### Credentials

- VM admin password: Stored securely (use Key Vault in production)
- GitHub token: Expires after 1 hour (runner uses it once during setup)

### Best Practices

1. Change NSG to allow RDP only from your IP
2. Use Azure Bastion for RDP access (no public IP)
3. Enable Azure Security Center
4. Regular Windows Updates
5. Rotate VM admin password quarterly

## Support

- **Azure Issues:** Azure Support Portal
- **GitHub Runner:** [GitHub Docs](https://docs.github.com/en/actions/hosting-your-own-runners)
- **Repository Issues:** https://github.com/sbroenne/mcp-server-excel/issues

## Cleanup

To delete all resources:

```bash
az group delete --name rg-excel-runner --yes --no-wait
```

This removes VM, disks, network resources, and stops all charges.

---

**Cost Estimate:** ~$30/month with auto-shutdown  
**Setup Time:** 5 min deploy + 30 min Excel install  
**Maintenance:** ~15 min/month
