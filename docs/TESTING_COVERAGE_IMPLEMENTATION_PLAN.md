# Implementation Plan: Improve Testing Coverage with Azure Self-Hosted Runners

**Status:** 📋 Implementation Plan  
**Created:** 2025-10-31  
**Issue:** [FEATURE] Improve testing coverage

## Problem Statement

ExcelMcp requires Microsoft Excel for integration testing via COM automation. Current CI/CD runs on GitHub-hosted `windows-latest` runners which **do not include Microsoft Excel**. This means:

- ❌ Integration tests (Category=Integration) are **skipped** in CI/CD
- ✅ Only Unit tests (Category=Unit) run in CI/CD
- ⚠️ Integration test coverage exists but requires manual local execution
- 📊 ~91 integration tests exist but don't run automatically

## Solution: Azure Self-Hosted Runners

Deploy a Windows VM in Azure with Microsoft Excel and GitHub Actions self-hosted runner to enable automated integration testing.

### Architecture Overview

```
GitHub Actions Workflow
    ↓ (trigger)
Azure Windows VM
    - Windows Server 2022
    - .NET 8 SDK
    - Microsoft Excel (Office 365)
    - GitHub Actions Runner
    ↓ (execute)
Integration Tests (Category=Integration)
    - Core.Tests (Excel COM operations)
    - CLI.Tests (Command-line interface)
    - McpServer.Tests (MCP protocol)
    - ComInterop.Tests (COM session management)
```

## Implementation Approach

### ✅ Phase 1: Documentation (Current)

**Deliverables:**
1. **`docs/AZURE_SELFHOSTED_RUNNER_SETUP.md`** - Complete setup guide
   - Azure VM provisioning (Portal + CLI)
   - .NET 8 SDK installation
   - Microsoft Excel installation & configuration
   - GitHub runner installation & service setup
   - Network security & firewall configuration
   - Cost estimates & optimization strategies
   - Troubleshooting & maintenance procedures

2. **`.github/workflows/integration-tests.yml`** - New workflow
   - Runs on `[self-hosted, windows, excel]` labels
   - Scheduled nightly execution (2 AM UTC)
   - Manual workflow_dispatch trigger
   - Executes all integration tests (Category=Integration)
   - Excel process cleanup after tests
   - Test result artifacts uploaded

### 🔄 Phase 2: Runner Deployment (To Be Done)

**Prerequisites:**
- Azure subscription with VM creation permissions
- Office 365 license (E3/E5 or standalone Excel)
- GitHub repository admin access

**Tasks:**
1. Provision Azure Windows VM (Standard_D2s_v3 recommended)
2. Install .NET 8 SDK on VM
3. Install & activate Microsoft Excel
4. Configure Excel for automation (disable splash screens, enable VBA access)
5. Install GitHub Actions runner as Windows service
6. Register runner with `self-hosted`, `windows`, `excel` labels
7. Configure auto-shutdown schedule to save costs
8. Test runner with integration-tests.yml workflow

### 🚀 Phase 3: CI/CD Integration (Optional)

**Tasks:**
1. Add integration test status badge to README.md
2. Update existing workflows to reference integration tests
3. Configure notifications for integration test failures
4. Set up monitoring for runner health

## Cost Analysis

### Monthly Cost (Pay-as-you-go, East US)

| Resource | Specification | Cost/Month |
|----------|--------------|------------|
| VM (Standard_D2s_v3) | 2 vCPUs, 8 GB RAM | ~$70 |
| Storage (Premium SSD) | 128 GB P10 | ~$20 |
| Network | ~10 GB egress | ~$1 |
| **Total (24/7)** | | **~$91** |

### Cost Optimization

1. **Auto-shutdown** (7 PM daily)
   - Running 12 hours/day: **~$45/month**
   - Storage always charged: **+$20**
   - **Total: ~$65/month**

2. **Scheduled runs only**
   - Start VM for nightly tests
   - Run for 2 hours/day: **~$10/month**
   - Storage always charged: **+$20**
   - **Total: ~$30/month**

3. **B-series burstable VM**
   - Standard_B2s: 2 vCPUs, 4 GB RAM
   - Running 24/7: **~$30/month**
   - Storage: **+$20**
   - **Total: ~$50/month**

**Recommended:** Start with auto-shutdown schedule (~$65/month), move to scheduled-only if feasible (~$30/month).

## Testing Strategy

### Current State

```bash
# CI/CD (GitHub-hosted runners) - No Excel
dotnet test --filter "Category=Unit&RunType!=OnDemand"
# Result: ~46 unit tests pass

# Local Developer (with Excel installed)
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand"
# Result: ~137 tests pass (46 unit + 91 integration)

# Pool cleanup (manual, with Excel)
dotnet test --filter "RunType=OnDemand"
# Result: ~5 tests (Excel process cleanup validation)
```

### Future State (with Azure runner)

```bash
# CI/CD Unit Tests (GitHub-hosted runners) - No change
dotnet test --filter "Category=Unit&RunType!=OnDemand"
# Result: ~46 unit tests

# CI/CD Integration Tests (Self-hosted Azure runner) - NEW
dotnet test --filter "Category=Integration&RunType!=OnDemand"
# Result: ~91 integration tests

# Total automated coverage: ~137 tests
```

### Test Categories

| Category | Tests | Requires Excel | Runs In | Execution Time |
|----------|-------|----------------|---------|----------------|
| Unit | ~46 | ❌ No | GitHub-hosted | 2-5 sec |
| Integration | ~91 | ✅ Yes | Self-hosted Azure | 13-15 min |
| OnDemand | ~5 | ✅ Yes | Manual only | 3-5 min |
| **Total** | **~142** | | | |

## Alternatives Considered

### ❌ GitHub-Hosted Larger Runners with Excel

**Why rejected:** GitHub-hosted runners (even enterprise larger runners) do not include Microsoft Office/Excel.

### ❌ Third-Party CI/CD with Excel

**Options:** BuildJet, Cirrus CI  
**Why rejected:** Higher cost, vendor lock-in, less control over environment.

### ❌ Docker/Containers with Excel

**Why rejected:** Excel COM automation requires full Windows Desktop experience. Windows Server Core containers don't support Excel UI automation.

### ❌ Mock/Stub Excel COM

**Why rejected:** ExcelMcp's core value proposition is **real Excel COM automation**. Mocking would not test actual Excel behavior (Power Query engine, VBA runtime, Data Model, calculation engine).

### ✅ Azure Self-Hosted Runner (Selected)

**Advantages:**
- Full control over environment
- Real Excel COM testing
- Flexible cost management (auto-shutdown, scheduled runs)
- Integrates seamlessly with GitHub Actions
- Can scale with VMSS in future

**Disadvantages:**
- Monthly Azure costs (~$30-$91)
- Requires VM maintenance
- Runner configuration overhead

## Implementation Checklist

- [x] Create comprehensive setup documentation
- [x] Create integration test workflow
- [x] Document cost analysis
- [x] Document testing strategy
- [ ] Provision Azure Windows VM (requires Azure access)
- [ ] Install .NET 8 SDK on VM
- [ ] Install & configure Microsoft Excel
- [ ] Install GitHub Actions runner
- [ ] Register runner with repository
- [ ] Test integration-tests.yml workflow
- [ ] Configure auto-shutdown
- [ ] Update README with integration test badge
- [ ] Set up monitoring/alerts

## Success Criteria

1. ✅ Documentation complete and comprehensive
2. ⏳ Azure runner deployed and registered
3. ⏳ Integration tests run successfully on self-hosted runner
4. ⏳ Nightly scheduled runs execute without intervention
5. ⏳ Cost stays under $50/month with optimizations
6. ⏳ Runner maintenance documented and sustainable

## Future Enhancements

### Potential Improvements

1. **Terraform/IaC Automation**
   - Automate VM provisioning with Terraform
   - Infrastructure as Code for reproducibility
   - Easy disaster recovery

2. **Azure VM Scale Sets (VMSS)**
   - Auto-scale runners based on workload
   - Ephemeral runners for cost optimization
   - Parallel test execution

3. **Azure DevTest Labs**
   - Pre-configured VM images with Excel
   - Fast provisioning from snapshots
   - Better cost management

4. **Hybrid Approach**
   - Keep unit tests on GitHub-hosted runners
   - Only integration tests on self-hosted
   - Parallel execution for faster feedback

5. **Multi-Region Runners**
   - Deploy runners in multiple Azure regions
   - Faster access for distributed teams
   - Geographic redundancy

## References

- [GitHub Self-Hosted Runners Documentation](https://docs.github.com/en/actions/hosting-your-own-runners)
- [Azure Virtual Machines Pricing](https://azure.microsoft.com/en-us/pricing/details/virtual-machines/windows/)
- [Office Deployment Tool](https://learn.microsoft.com/en-us/deployoffice/overview-office-deployment-tool)
- [GitHub Actions Best Practices](https://docs.github.com/en/actions/security-guides/security-hardening-for-github-actions)

## Next Steps

**For Repository Maintainer:**
1. Review implementation plan
2. Approve Azure budget (~$30-$65/month)
3. Provide Azure subscription access
4. Provide Office 365 license
5. Follow setup guide in `docs/AZURE_SELFHOSTED_RUNNER_SETUP.md`
6. Test integration workflow
7. Configure auto-shutdown schedule
8. Monitor costs and runner health

**Timeline Estimate:**
- Initial setup: 2-4 hours
- Testing & validation: 1-2 hours
- Documentation updates: 30 minutes
- **Total: 4-7 hours** (one-time setup)

---

**Status:** ✅ Documentation complete, awaiting Phase 2 deployment approval
