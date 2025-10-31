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

| Resource | Type | Purpose | Monthly Cost (Sweden Central) |
|----------|------|---------|-------------------------------|
| VM | Standard_B2ms (2 vCPUs, 8 GB RAM) | Test runner | ~$50 (24/7) or ~$25 (12h/day) |
| OS Disk | Premium SSD 128 GB | Storage | ~$11 |
| Network | VNet, NIC, NSG, Public IP | Connectivity | <$1 |
| **Total (24/7)** | | | **~$61/month** |
| **Total (12h/day)** | | | **~$36/month** |

### ⚠️ Important: GitHub Actions Cannot Auto-Start VMs

**GitHub Actions runners require the VM to be running.** If the VM is stopped:
- Workflows will queue and wait indefinitely
- No automatic VM start capability exists

**You have 3 options:**

1. **Keep VM running 24/7** (~$61/month) ⭐ **RECOMMENDED for active development**
   - Workflows execute immediately
   - No automation needed
   - Simplest setup

2. **Manual start/stop** (~$36/month with 12h/day)
   - Start VM manually each morning
   - Auto-shutdown at 7 PM
   - Good for predictable schedules

3. **Azure Automation** (~$36/month with scheduled start)
   - Use Azure Automation Runbook to start VM at specific times
   - Auto-shutdown at night
   - Requires additional setup (see below)

### Location & VM Size

- **Location:** Sweden Central (your preference)
- **VM Size:** Standard_B2ms (2 vCPUs, 8 GB RAM)
  - 8GB RAM required for reliable Excel automation
  - Burstable performance for cost efficiency
- **Auto-shutdown:** 7 PM UTC daily (optional, saves ~40%)

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

| Size | vCPUs | RAM | Monthly Cost (Sweden Central 24/7) | Use Case |
|------|-------|-----|-----------------------------------|----------|
| Standard_B2s | 2 | 4 GB | ~$40 | Too small for Excel |
| **Standard_B2ms** | **2** | **8 GB** | **~$61** | **Recommended** ⭐ |
| Standard_B4ms | 4 | 16 GB | ~$120 | Overkill for testing |

**Note:** 8GB RAM is required for reliable Excel COM automation with multiple test projects.

### Region Options

| Region | Monthly Cost (B2ms, 24/7) | Latency to GitHub | Notes |
|--------|---------------------------|-------------------|-------|
| East US | ~$51 | Low | Cheapest |
| West Europe | ~$58 | Medium | EU data residency |
| **Sweden Central** | **~$61** | **Medium** | **Your choice** ⭐ |
| North Europe | ~$58 | Medium | EU alternative |

**Selected:** Sweden Central (as per your preference)

### Auto-Shutdown Schedule

Default: 7 PM UTC (19:00)

Edit in `azure-runner.bicep`:
```bicep
dailyRecurrence: {
  time: '1900' // Change to '2000' for 8 PM, etc.
}
```

## Cost Optimization

### Recommended: Keep VM Running 24/7 (~$61/month)

**Why:** GitHub Actions workflows execute immediately when code changes are pushed.

**Cost:** ~$61/month in Sweden Central

**Best for:** Active development with frequent commits

---

### Alternative: Auto-Start with Azure Automation (~$36/month)

If you want to save costs but maintain automation, use Azure Automation:

**Setup Azure Automation (one-time):**

```bash
# Create Automation Account
az automation account create \
  --name "automation-excel-runner" \
  --resource-group "rg-excel-runner" \
  --location "swedencentral"

# Create Start-VM Runbook
# Upload PowerShell script to start VM at 7 AM daily
```

**PowerShell Runbook (StartVM.ps1):**
```powershell
param(
    [Parameter(Mandatory=$true)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory=$true)]
    [string]$VMName
)

# Authenticate with Managed Identity
Connect-AzAccount -Identity

# Start VM
Start-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName
```

**Schedule:**
- Start: 7 AM UTC (Monday-Friday)
- Stop: 7 PM UTC (auto-shutdown configured in template)

**Cost Breakdown:**
- VM (12h/day, weekdays): ~$25/month
- Storage (always): ~$11/month
- Automation: ~$0.50/month
- **Total: ~$36/month**

**Tradeoff:** Workflows pushed outside 7 AM - 7 PM will queue until VM starts.

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
