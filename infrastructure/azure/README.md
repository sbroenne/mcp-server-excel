# Azure Infrastructure for Excel Integration Testing

This directory contains Infrastructure as Code (IaC) for automating the deployment of Azure Windows VM with GitHub Actions self-hosted runner.

## Quick Start

**Deploy directly from GitHub UI - fully automated!**

📚 **Setup Guide:** [`GITHUB_ACTIONS_DEPLOYMENT.md`](GITHUB_ACTIONS_DEPLOYMENT.md)

**Quick steps:**
1. Create Azure App Registration with OIDC (one-time, 10 minutes)
2. Add Azure credentials to GitHub Secrets  
3. Go to Actions tab → Deploy Azure Self-Hosted Runner
4. Enter parameters (Resource Group + Admin Password only - **no manual token needed!**)
5. RDP to VM and install Excel (30 minutes)

**Benefits:**
- ✅ **Fully automated** - runner token auto-generated via GitHub CLI (secure & reliable)
- ✅ **No local tooling required** - deploy from browser
- ✅ **Secure OIDC authentication** - no stored secrets
- ✅ **Audit trail** in Actions logs
- ✅ **Repeatable and version-controlled**

---

## What Gets Deployed

### Resources Created

| Resource | Type | Purpose | Monthly Cost (Sweden Central) |
|----------|------|---------|-------------------------------|
| VM | Standard_B2ms (2 vCPUs, 8 GB RAM) | Test runner | ~$50 |
| OS Disk | Premium SSD 128 GB | Storage | ~$11 |
| Network | VNet, NIC, NSG, Public IP | Connectivity | <$1 |
| **Total (24/7)** | | | **~$61/month** |

### GitHub Coding Agent Compatibility

**✅ YES** - GitHub Coding Agents can use this runner in Agent mode:

1. **Runner labels:** `self-hosted`, `windows`, `excel`
2. **Available 24/7** for immediate workflow execution
3. **Coding agents access runner same way as workflows:**
   - When you push code in VS Code Agent mode
   - Workflows using `runs-on: [self-hosted, windows, excel]`
   - Manual workflow triggers via Actions tab

**How it works:**
- GitHub Coding Agent pushes commits → triggers workflow → runs on your self-hosted runner
- No difference between coding agent commits and manual commits
- Runner executes tests automatically on every push

### Location & VM Size

- **Location:** Sweden Central
- **VM Size:** Standard_B2ms (2 vCPUs, 8 GB RAM)
  - 8GB RAM required for reliable Excel automation
  - Burstable performance for cost efficiency
- **Uptime:** 24/7 (VM runs continuously for immediate test execution)

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
├── azure-runner.bicep                # Main Bicep template
├── GITHUB_ACTIONS_DEPLOYMENT.md      # Setup and deployment guide
└── README.md                         # This file
```

## Configuration

### VM Size Options

| Size | vCPUs | RAM | Monthly Cost (Sweden Central 24/7) | Use Case |
|------|-------|-----|-----------------------------------|----------|
| Standard_B2s | 2 | 4 GB | ~$40 | Too small for Excel |
| **Standard_B2ms** | **2** | **8 GB** | **~$61** | **Recommended** ⭐ |
| Standard_B4ms | 4 | 16 GB | ~$120 | Overkill for testing |

**Note:** 8GB RAM is required for reliable Excel COM automation with multiple test projects.

### Region Options

| Region | Monthly Cost (B2ms, 24/7) | Notes |
|--------|---------------------------|-------|
| East US | ~$51 | Cheapest |
| West Europe | ~$58 | EU data residency |
| **Sweden Central** | **~$61** | **Selected** ⭐ |
| North Europe | ~$58 | EU alternative |

**Cost:** ~$61/month for 24/7 operation in Sweden Central

### Auto-Shutdown Schedule

**Removed** - Auto-shutdown disabled for 24/7 availability.

VM runs continuously to ensure:
- Immediate workflow execution on every commit
- No queued workflows waiting for VM start
- Best experience for GitHub Coding Agents
- Simplified operation (no manual VM starts)

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
