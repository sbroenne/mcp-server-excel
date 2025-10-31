# Quick Start: Enable Integration Testing

> **For Repository Owner** - 30-minute setup to enable Excel COM integration testing in CI/CD

## What This Gives You

- ✅ **Automated Excel testing** in CI/CD (91 integration tests)
- ✅ **Nightly test runs** catch regressions early
- ✅ **Manual test triggers** for on-demand validation
- 💰 **Estimated cost:** $30-65/month (with optimizations)

## Prerequisites Checklist

- [ ] Azure subscription with VM creation permissions
- [ ] Office 365 license (E3/E5 or standalone Excel)
- [ ] GitHub repository admin access
- [ ] ~4-7 hours for setup (one-time)

## Setup Steps (High-Level)

### 1. Create Azure VM (30 minutes)
```powershell
# Quick Azure CLI setup
az login
az group create --name rg-excel-runner --location eastus
az vm create \
  --resource-group rg-excel-runner \
  --name vm-excel-runner-01 \
  --image Win2022Datacenter \
  --size Standard_D2s_v3 \
  --admin-username adminuser \
  --admin-password 'YourSecurePassword!'
```

📚 **Full instructions:** [docs/AZURE_SELFHOSTED_RUNNER_SETUP.md](AZURE_SELFHOSTED_RUNNER_SETUP.md) (Step 1)

### 2. Connect & Install Prerequisites (1 hour)

**RDP into VM, then run:**
```powershell
# Install .NET 8
Invoke-WebRequest -Uri https://aka.ms/dotnet/8.0/dotnet-sdk-win-x64.exe -OutFile dotnet-sdk.exe
Start-Process dotnet-sdk.exe -ArgumentList '/quiet' -Wait

# Install Office/Excel (requires Office 365 account)
# See full guide for Office Deployment Tool instructions
```

📚 **Full instructions:** [docs/AZURE_SELFHOSTED_RUNNER_SETUP.md](AZURE_SELFHOSTED_RUNNER_SETUP.md) (Step 2)

### 3. Install GitHub Runner (30 minutes)

**On VM:**
1. Go to GitHub: `Settings` → `Actions` → `Runners` → `New self-hosted runner`
2. Copy registration token
3. Run on VM:
   ```powershell
   New-Item -Path "C:\actions-runner" -ItemType Directory
   Set-Location "C:\actions-runner"
   # Download runner (see full guide for latest version)
   # Configure with token
   # Install as Windows service
   ```

📚 **Full instructions:** [docs/AZURE_SELFHOSTED_RUNNER_SETUP.md](AZURE_SELFHOSTED_RUNNER_SETUP.md) (Step 3)

### 4. Test Integration Workflow (15 minutes)

**In GitHub:**
1. Go to `Actions` tab
2. Select `Integration Tests (Excel)` workflow
3. Click `Run workflow` → `Run workflow`
4. Wait for completion (~15 minutes)
5. Verify ✅ all tests pass

📚 **Full instructions:** [docs/AZURE_SELFHOSTED_RUNNER_SETUP.md](AZURE_SELFHOSTED_RUNNER_SETUP.md) (Step 6)

### 5. Configure Cost Savings (15 minutes)

**In Azure Portal:**
1. Go to VM → `Auto-shutdown`
2. Enable: `On`
3. Time: `19:00` (7 PM)
4. **Save**

**Result:** ~$65/month (vs ~$91/month 24/7)

📚 **Full instructions:** [docs/AZURE_SELFHOSTED_RUNNER_SETUP.md](AZURE_SELFHOSTED_RUNNER_SETUP.md) (Maintenance section)

## Cost Breakdown

| Operation Mode | Monthly Cost | When to Use |
|----------------|--------------|-------------|
| 24/7 | ~$91 | Always available for tests |
| Auto-shutdown (12h/day) | ~$65 | Standard setup ⭐ |
| Scheduled only (2h/day) | ~$30 | Budget-conscious |

💡 **Recommendation:** Start with auto-shutdown, optimize later if needed.

## What Gets Tested

**Before (GitHub-hosted runners):**
- ✅ 46 unit tests (no Excel required)
- ❌ 91 integration tests **SKIPPED**

**After (with Azure runner):**
- ✅ 46 unit tests (GitHub-hosted)
- ✅ 91 integration tests (Azure runner) **← NEW**
- **Total: 137 automated tests** (96% coverage)

## Monitoring & Maintenance

**Check runner status:**
- GitHub: `Settings` → `Actions` → `Runners`
- Should show: `azure-excel-runner` (Idle/Active)

**Monthly tasks:**
- Update runner agent (notifications in GitHub)
- Windows Updates on VM
- Office Updates

**Cost monitoring:**
- Azure Portal: `Cost Management + Billing`
- Set budget alerts at $50, $75, $100

## Troubleshooting Quick Links

| Issue | Solution |
|-------|----------|
| Runner offline | [Troubleshooting Guide](AZURE_SELFHOSTED_RUNNER_SETUP.md#runner-shows-offline) |
| Excel COM errors | [Excel Troubleshooting](AZURE_SELFHOSTED_RUNNER_SETUP.md#excel-com-errors-in-tests) |
| Tests timeout | [Timeout Guide](AZURE_SELFHOSTED_RUNNER_SETUP.md#tests-timeout) |
| High costs | [Cost Optimization](AZURE_SELFHOSTED_RUNNER_SETUP.md#cost-optimization-strategies) |

## Support & Documentation

📖 **Complete guides:**
- [Full Setup Guide](AZURE_SELFHOSTED_RUNNER_SETUP.md) - Step-by-step instructions
- [Implementation Plan](TESTING_COVERAGE_IMPLEMENTATION_PLAN.md) - Architecture & rationale

🐛 **Issues:**
- GitHub Issues: https://github.com/sbroenne/mcp-server-excel/issues

## Success Checklist

After setup, verify:

- [ ] VM running in Azure Portal
- [ ] Runner shows "Idle" in GitHub Settings → Actions → Runners
- [ ] Integration Tests workflow runs successfully
- [ ] Auto-shutdown configured (saves ~$26/month)
- [ ] Cost alerts set up in Azure
- [ ] Team knows how to trigger manual test runs

---

**Ready to start?** → [Full Setup Guide](AZURE_SELFHOSTED_RUNNER_SETUP.md)
