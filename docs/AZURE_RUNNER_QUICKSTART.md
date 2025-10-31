# Azure Runner Setup - Quick Reference

This is a quick reference for setting up the Azure self-hosted runner for Excel integration testing.

## 🚀 Quick Decision Tree

```
Do you have an Azure Windows VM already?
│
├─ NO → Use automated deployment
│   └─ Go to: infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md
│
└─ YES → Use manual installation
    └─ Go to: docs/MANUAL_RUNNER_INSTALLATION.md
```

## 📚 Documentation Map

| Scenario | Document | Time |
|----------|----------|------|
| **First-time setup (no VM)** | [GITHUB_ACTIONS_DEPLOYMENT.md](../infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md) | 5 min + 30 min Excel |
| **Manual installation (existing VM)** | [MANUAL_RUNNER_INSTALLATION.md](MANUAL_RUNNER_INSTALLATION.md) | 15 min + 30 min Excel |
| **Automated deployment failed** | [MANUAL_RUNNER_INSTALLATION.md](MANUAL_RUNNER_INSTALLATION.md) | 15 min + 30 min Excel |
| **Infrastructure overview** | [infrastructure/azure/README.md](../infrastructure/azure/README.md) | Read only |
| **General Azure setup info** | [AZURE_SELFHOSTED_RUNNER_SETUP.md](AZURE_SELFHOSTED_RUNNER_SETUP.md) | Read only |

## ⚡ Super Quick Start (Automated)

**Prerequisites:** Azure subscription, Office 365 license

1. **Setup Azure OIDC** (one-time, 10 min):
   ```bash
   # See: infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md
   # Create App Registration with federated credentials
   # Add AZURE_CLIENT_ID, AZURE_TENANT_ID, AZURE_SUBSCRIPTION_ID to GitHub Secrets
   ```

2. **Get Runner Token** (expires in 1 hour):
   - Go to: https://github.com/sbroenne/mcp-server-excel/settings/actions/runners/new
   - Copy the token (starts with 'A')

3. **Deploy via GitHub Actions**:
   - Go to: Actions → Deploy Azure Self-Hosted Runner
   - Enter: Resource Group, Admin Password, Runner Token
   - Click: Run workflow
   - Wait: 5 minutes

4. **Install Excel**:
   - RDP to VM (check workflow output for IP)
   - Install Excel from portal.office.com
   - Activate Excel
   - Reboot

**Done!** Verify at: https://github.com/sbroenne/mcp-server-excel/settings/actions/runners

## 🔧 Manual Installation (Existing VM)

**When to use:** Automated deployment failed OR you have existing VM

**Steps** (see [MANUAL_RUNNER_INSTALLATION.md](MANUAL_RUNNER_INSTALLATION.md) for details):

1. RDP to VM
2. Install .NET 8 SDK
3. Generate GitHub runner token (expires in 1 hour)
4. Download and configure GitHub Actions runner
5. Install as Windows service
6. Install Office 365 Excel
7. Verify Excel COM access
8. Verify runner registration

**Time:** 15 minutes + 30 minutes for Excel

## ❓ Troubleshooting

| Problem | Solution | Document |
|---------|----------|----------|
| Automated deployment failed | Check VM logs at `C:\runner-setup.log` via RDP | [GITHUB_ACTIONS_DEPLOYMENT.md](../infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md#troubleshooting) |
| Runner token expired | Generate new token, expires in 1 hour | [MANUAL_RUNNER_INSTALLATION.md](MANUAL_RUNNER_INSTALLATION.md#troubleshooting) |
| Runner service won't start | Check service logs via `Get-EventLog` | [MANUAL_RUNNER_INSTALLATION.md](MANUAL_RUNNER_INSTALLATION.md#troubleshooting) |
| Excel COM test fails | Install/activate Excel, kill background processes | [MANUAL_RUNNER_INSTALLATION.md](MANUAL_RUNNER_INSTALLATION.md#troubleshooting) |
| Workflow not using runner | Check `runs-on: [self-hosted, windows, excel]` | [MANUAL_RUNNER_INSTALLATION.md](MANUAL_RUNNER_INSTALLATION.md#troubleshooting) |

## 💰 Cost

**Standard_B2ms in Sweden Central (8GB RAM):**
- VM: ~$50/month
- Storage: ~$11/month
- Network: <$1/month
- **Total: ~$61/month** (24/7 operation)

**Why 24/7?**
- Immediate CI/CD execution
- Best experience for GitHub Coding Agents
- No queued workflows waiting for VM start

## 🔒 Security

- Use OIDC for Azure authentication (no client secrets)
- Runner tokens expire after 1 hour
- Restrict RDP via NSG (configure your IP)
- Runner runs as Network Service (least privilege)
- Enable Windows Defender and auto-updates

## 📊 Verification Checklist

After setup, verify:

- [ ] Runner shows in GitHub: https://github.com/sbroenne/mcp-server-excel/settings/actions/runners
- [ ] Runner status: Idle (green)
- [ ] Runner labels: `self-hosted`, `windows`, `excel`
- [ ] Excel COM test passes: `New-Object -ComObject Excel.Application`
- [ ] Integration tests workflow runs successfully
- [ ] Cost alerts configured: $40/month threshold

## 🆘 Support

- **Automated deployment issues:** Check workflow logs in Actions tab
- **Manual installation issues:** See [MANUAL_RUNNER_INSTALLATION.md](MANUAL_RUNNER_INSTALLATION.md)
- **General questions:** [Create Issue](https://github.com/sbroenne/mcp-server-excel/issues/new)

---

**Last Updated:** 2025-10-31  
**Recommended Approach:** Automated deployment with manual fallback if needed
