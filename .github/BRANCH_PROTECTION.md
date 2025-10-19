# Branch Protection Configuration Guide

This document provides step-by-step instructions for setting up branch protection rules for the ExcelMcp repository to ensure code quality, security, and proper PR-based workflow.

## Prerequisites

- Repository administrator access
- GitHub Pro, Team, or Enterprise account (for some advanced features)

## Main Branch Protection Rules

### Step 1: Navigate to Branch Protection Settings

1. Go to your repository on GitHub
2. Click **Settings** (gear icon)
3. Click **Branches** in the left sidebar
4. Click **Add branch protection rule**
5. Enter `main` as the branch name pattern

### Step 2: Configure Protection Rules

Apply the following settings for the `main` branch:

#### ✅ Require Pull Request Reviews Before Merging

**Enable:** ✓ Require a pull request before merging

**Settings:**
- ✓ Require approvals: **1** (minimum)
- ✓ Dismiss stale pull request approvals when new commits are pushed
- ✓ Require review from Code Owners (optional - requires CODEOWNERS file)
- ✓ Require approval of the most recent reviewable push

**Purpose:** Ensures all code is reviewed before merging, catching issues early.

#### ✅ Require Status Checks to Pass Before Merging

**Enable:** ✓ Require status checks to pass before merging

**Settings:**
- ✓ Require branches to be up to date before merging
- **Required status checks** (add these):
  - `build` - Build workflow must pass
  - `analyze / Analyze Code with CodeQL` - CodeQL security scan must pass
  - `dependency-review` - Dependency review must pass
  
**Purpose:** Ensures code builds successfully and passes security scans before merging.

**Note:** Status checks will only appear after they've run at least once. Create a test PR to see them populate.

#### ✅ Require Conversation Resolution Before Merging

**Enable:** ✓ Require conversation resolution before merging

**Purpose:** Ensures all review comments are addressed before merging.

#### ✅ Require Signed Commits (Recommended)

**Enable:** ✓ Require signed commits

**Purpose:** Ensures all commits are cryptographically verified (GPG/SSH signature).

**Setup Guide:** https://docs.github.com/en/authentication/managing-commit-signature-verification

#### ✅ Require Linear History (Optional)

**Enable:** ✓ Require linear history

**Purpose:** Prevents merge commits, keeps history clean with rebase/squash only.

**Note:** This may conflict with some workflows. Enable based on team preference.

#### ✅ Include Administrators

**Disable:** ☐ Do not allow bypassing the above settings (even for admins)

**Purpose:** Enforces rules for everyone, including repository administrators.

**Recommendation:** Enable this for maximum security, disable if you need emergency access.

#### ✅ Restrict Who Can Push to Matching Branches

**Enable:** ✓ Restrict pushes that create matching branches

**Settings:**
- Add specific users or teams who can push (usually just CI/CD service accounts)
- Most developers should only merge through PRs

**Purpose:** Prevents direct pushes to main branch.

#### ✅ Allow Force Pushes

**Disable:** ☐ Allow force pushes (keep unchecked)

**Purpose:** Prevents history rewriting on main branch.

#### ✅ Allow Deletions

**Disable:** ☐ Allow deletions (keep unchecked)

**Purpose:** Prevents accidental deletion of main branch.

### Step 3: Save Protection Rules

Click **Create** or **Save changes** at the bottom of the page.

## Develop Branch Protection Rules (Optional)

For the `develop` branch, you can use slightly relaxed rules:

1. Create another branch protection rule for `develop`
2. Apply similar settings but with:
   - Require approvals: **1** (same as main)
   - All status checks required (same as main)
   - Optional: Disable "Require branches to be up to date" for faster iteration
   - Optional: Allow force pushes (for rebasing feature branches)

## Feature Branch Naming Convention

Recommended branch naming patterns:
- `feature/*` - New features
- `bugfix/*` - Bug fixes
- `hotfix/*` - Critical production fixes
- `release/*` - Release preparation
- `copilot/*` - AI-generated changes (current pattern)

## PR Workflow with Branch Protection

### For Contributors

1. **Create Feature Branch**
   ```bash
   git checkout -b feature/my-feature
   ```

2. **Make Changes and Commit**
   ```bash
   git add .
   git commit -m "feat: add new feature"
   git push origin feature/my-feature
   ```

3. **Create Pull Request**
   - Go to GitHub and create PR from your branch to `main`
   - Fill out PR template
   - Request review from team members

4. **Address Review Comments**
   - Make requested changes
   - Push new commits to the same branch
   - Mark conversations as resolved

5. **Wait for Status Checks**
   - Build must pass
   - CodeQL must pass
   - Dependency review must pass (if dependencies changed)

6. **Merge PR**
   - Once approved and all checks pass, merge using:
     - **Squash and merge** (recommended - keeps history clean)
     - **Rebase and merge** (if linear history required)
     - **Create a merge commit** (preserves feature branch history)

### For Reviewers

1. **Review Code Changes**
   - Check code quality and logic
   - Verify tests are included
   - Check for security issues

2. **Check Status Checks**
   - Ensure all automated checks pass
   - Review CodeQL findings if any

3. **Approve or Request Changes**
   - Use GitHub's review feature
   - Provide constructive feedback
   - Request changes if needed

4. **Merge After Approval**
   - Wait for all required approvals
   - Ensure conversations are resolved
   - Merge using preferred method

## Automation with GitHub Actions

With branch protection enabled, the CI/CD pipeline works as follows:

### On Pull Request

1. **Build Workflows** (`build-mcp-server.yml`, `build-cli.yml`)
   - Triggers on: PR to main with code changes
   - Runs: Separate build and verification for MCP Server and CLI
   - Output: Build artifacts for both components
   - Duration: ~2-3 minutes each

2. **CodeQL Analysis** (`codeql.yml`)
   - Triggers on: PR to main with code changes
   - Runs: Security code scanning
   - Output: Security findings
   - Duration: ~5-10 minutes

3. **Dependency Review** (`dependency-review.yml`)
   - Triggers on: PR to main (always)
   - Runs: Vulnerability and license checking
   - Output: Dependency analysis
   - Duration: ~1 minute

### On Push to Main

1. **Build Workflow**
   - Builds and uploads artifacts
   - Validates deployment readiness

2. **CodeQL Analysis**
   - Runs full security scan
   - Updates security dashboard

3. **Dependabot**
   - Monitors for new vulnerabilities
   - Creates PRs for updates

### Weekly Schedule

1. **CodeQL** (Monday 10 AM UTC)
   - Full security scan
   - Updates security posture

2. **Dependabot** (Monday 9 AM EST)
   - Checks for dependency updates
   - Creates update PRs

## Path Filters for Efficient CI/CD

The workflows are now optimized to only run when relevant files change:

### Build-Relevant Changes
```yaml
paths:
  - 'src/**'              # Source code
  - 'tests/**'            # Test code
  - '**.csproj'           # Project files
  - '**.sln'              # Solution files
  - 'Directory.Build.props'    # Build configuration
  - 'Directory.Packages.props' # Package versions
  - '.github/workflows/build-*.yml'  # Build workflows
```

### Non-Build Changes (workflows skip)
- `README.md` and other documentation
- `docs/**` directory
- `.github/ISSUE_TEMPLATE/**`
- `SECURITY.md`
- License files

**Benefit:** Saves CI/CD minutes and provides faster feedback on documentation-only changes.

## Troubleshooting

### Problem: Required Status Checks Not Appearing

**Solution:**
1. Create a test PR
2. Let workflows run at least once
3. Return to branch protection settings
4. Status checks should now be selectable

### Problem: PR Cannot Be Merged Due to Outdated Branch

**Solution:**
```bash
git checkout feature/my-feature
git pull origin main
git push origin feature/my-feature
```

Or use GitHub's "Update branch" button on the PR.

### Problem: Status Check Failing

**Solution:**
1. Review the workflow logs on GitHub Actions tab
2. Fix the issue in your branch
3. Push the fix - checks will re-run automatically

### Problem: Cannot Force Push to Fix History

**Solution:**
This is intentional! Branch protection prevents force pushes to main. Instead:
1. Create a new commit with the fix
2. Or create a revert commit: `git revert <commit-hash>`

### Problem: Need to Merge Urgently

**Solution for Emergencies:**
1. Temporarily disable branch protection (admin only)
2. Make the critical change
3. Re-enable branch protection immediately
4. Document the reason in commit message

## Security Considerations

### Principle of Least Privilege
- Only grant repository admin access to team leads
- Most contributors should only have write access
- Use CODEOWNERS for sensitive files

### Audit Trail
- All changes require PR approval
- Signed commits provide verification
- GitHub maintains audit log of all actions

### Automated Security
- Dependabot alerts for vulnerabilities
- CodeQL scans for security issues
- Dependency review blocks risky dependencies

## Monitoring and Maintenance

### Weekly Tasks
1. Review Dependabot PRs
2. Check CodeQL findings
3. Review open PRs for stale reviews

### Monthly Tasks
1. Review branch protection rules
2. Update required status checks if workflows change
3. Audit team access levels

### Quarterly Tasks
1. Review security incidents
2. Update security policies
3. Train team on new security features

## Additional Resources

- [GitHub Branch Protection Documentation](https://docs.github.com/en/repositories/configuring-branches-and-merges-in-your-repository/managing-protected-branches/about-protected-branches)
- [Status Checks Documentation](https://docs.github.com/en/pull-requests/collaborating-with-pull-requests/collaborating-on-repositories-with-code-quality-features/about-status-checks)
- [Signed Commits Guide](https://docs.github.com/en/authentication/managing-commit-signature-verification)
- [CODEOWNERS Documentation](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/about-code-owners)

---

**Last Updated:** 2024-10-19

**Note:** These settings should be applied to the new `sbroenne/mcp-server-excel` repository once it's created.
