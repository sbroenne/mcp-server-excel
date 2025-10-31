# Quick Start: Enable Integration Testing

> **For Repository Owner** - Enable Excel COM integration testing in CI/CD

## Deployment Steps

**Time:** ~15 minutes setup + 30 minutes Excel installation  
**Automation:** GitHub Actions workflow handles everything automatically

### Step 1: Setup Azure OIDC (one-time, 10 minutes)

See [infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md](../infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md) for detailed instructions.

**Quick summary:**
1. Create Azure App Registration with OIDC federation
2. Assign Contributor role to your subscription
3. Add three secrets to GitHub repository:
   - `AZURE_CLIENT_ID`
   - `AZURE_TENANT_ID`
   - `AZURE_SUBSCRIPTION_ID`

### Step 2: Run Deployment Workflow (5 minutes)

1. Go to **Actions** tab → **Deploy Azure Self-Hosted Runner**
2. Click **Run workflow**
3. Enter:
   - **Resource Group:** `rg-excel-runner` (or your choice)
   - **Admin Password:** Strong password for VM access
4. Click **Run workflow**
5. Wait ~5 minutes for deployment

**✅ Runner token auto-generated** - no manual token handling required!

### Step 3: Install Excel (30 minutes)

1. RDP to VM (FQDN shown in workflow logs)
2. Sign in to https://portal.office.com
3. Install Office 365 Excel
4. Activate with your Office 365 account
5. Reboot VM

**✅ Done!** Runner auto-starts, integration tests enabled.

---

## What This Gives You

- ✅ **Automated Excel testing** in CI/CD (91 integration tests)
- ✅ **Nightly test runs** catch regressions early
- ✅ **Manual test triggers** for on-demand validation
- 💰 **Estimated cost:** $61/month (24/7 operation)

## Prerequisites Checklist

- [ ] Azure subscription with VM creation permissions
- [ ] Office 365 license (E3/E5 or standalone Excel)
- [ ] GitHub repository admin access

---

## Advanced Configuration (Optional)

### Cost Optimization

**VM runs 24/7 for immediate test execution.** If you want to reduce costs:

**Option 1: Auto-shutdown schedule**
1. Azure Portal → VM → Auto-shutdown
2. Enable and set shutdown time (e.g., 7 PM)
3. Saves ~$30/month but delays test runs

**Option 2: Start/stop on demand**
1. Stop VM when not actively developing
2. Start manually before pushing changes
3. Saves up to 50% but requires manual management

💡 **Recommendation:** Start with 24/7 operation for best developer experience.

### Test Integration Workflow

1. Go to `Actions` tab
2. Select `Integration Tests (Excel)` workflow
3. Click `Run workflow`
4. Wait ~15 minutes for completion
5. Verify ✅ all tests pass

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
- Set budget alerts at $50, $75

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Runner offline | Check VM is running in Azure Portal, restart runner service |
| Excel COM errors | Verify Excel is activated, check test logs |
| Tests timeout | Check VM resources, verify Excel not showing dialogs |
| Deployment fails | Check Azure credentials, verify OIDC setup |

## Support & Documentation

📖 **Complete guides:**
- [Deployment Guide](../infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md) - Detailed setup instructions
- [Implementation Plan](TESTING_COVERAGE_IMPLEMENTATION_PLAN.md) - Architecture & rationale

🐛 **Issues:**
- GitHub Issues: https://github.com/sbroenne/mcp-server-excel/issues

## Success Checklist

After setup, verify:

- [ ] VM running in Azure Portal
- [ ] Runner shows "Idle" in GitHub Settings → Actions → Runners
- [ ] Integration Tests workflow runs successfully
- [ ] Cost alerts set up in Azure

---

**Estimated time:** 15 min setup + 30 min Excel install = **45 minutes total**  
**Monthly cost:** ~$61 (24/7 operation)
