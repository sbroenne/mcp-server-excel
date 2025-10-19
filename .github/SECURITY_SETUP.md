# GitHub Advanced Security Setup Guide

This document describes the security features enabled for the ExcelMcp project.

## Enabled Security Features

### 1. Dependabot

### 0. CodeQL Configuration (Important - Read First!)

**IMPORTANT:** This repository includes an advanced CodeQL workflow configuration. You must choose between:

#### Option A: Use Advanced CodeQL Workflow (Recommended for this project)
This provides custom configuration with COM interop exceptions and path filters.

**Setup Steps:**
1. Go to Settings → Code security and analysis
2. If "CodeQL analysis" shows "Set up" → Click it and choose "Advanced"
3. If "CodeQL analysis" shows "Default" → Click "..." → "Switch to advanced"
4. Delete the auto-generated `.github/workflows/codeql-analysis.yml` if created
5. Keep our custom `.github/workflows/codeql.yml` 
6. The custom config `.github/codeql/codeql-config.yml` will be used automatically

**Why Advanced?**
- Custom path filters (only scans code changes, not docs)
- COM interop false positive exclusions
- Custom query suites (security-and-quality + security-extended)
- Windows runner for proper .NET/Excel COM build

#### Option B: Use Default Setup (Simpler but less customized)
GitHub manages everything automatically, but you lose custom configuration.

**Setup Steps:**
1. Delete `.github/workflows/codeql.yml`
2. Delete `.github/codeql/codeql-config.yml`
3. Go to Settings → Code security and analysis
4. Click "Set up" for CodeQL analysis → Choose "Default"

**Trade-offs:**
- ❌ No custom COM interop exclusions
- ❌ No path filters (scans on all changes)
- ❌ Cannot customize query suites
- ✅ Fully automated by GitHub
- ✅ No workflow maintenance

**Current Error Fix:**
If you see: `CodeQL analyses from advanced configurations cannot be processed when the default setup is enabled`

**Solution:** Follow Option A steps above to switch from default to advanced setup.

### 1. Dependabot
**Configuration**: `.github/dependabot.yml`

Automatically monitors and updates dependencies:
- **NuGet packages**: Weekly updates on Mondays
- **GitHub Actions**: Weekly updates on Mondays
- **Security alerts**: Immediate notifications for vulnerabilities

**Features**:
- Grouped updates for minor/patch versions
- Automatic PR creation for security updates
- License compliance checking
- Vulnerability scanning

**To enable** (Repository Admin only):
1. Go to Settings → Code security and analysis
2. Enable "Dependabot alerts"
3. Enable "Dependabot security updates"
4. Enable "Dependabot version updates"

### 2. CodeQL Analysis
**Configuration**: `.github/workflows/codeql.yml` and `.github/codeql/codeql-config.yml`

Performs advanced security code scanning:
- **Languages**: C#
- **Schedule**: Weekly on Mondays + every push/PR
- **Query suites**: security-and-quality, security-extended
- **ML-powered**: Advanced vulnerability detection

**Security Checks**:
- SQL injection (CA2100)
- Path traversal (CA3003)
- Command injection (CA3006)
- Cryptographic issues (CA5389, CA5390, CA5394)
- Input validation
- Resource management
- Sensitive data exposure

**To enable** (Repository Admin only):
1. Go to Settings → Code security and analysis
2. Enable "Code scanning"
3. Click "Set up" → "Advanced"
4. The workflow file is already configured

### 3. Dependency Review
**Configuration**: `.github/workflows/dependency-review.yml`

Reviews pull requests for:
- Vulnerable dependencies
- License compliance issues
- Breaking changes in dependencies

**Blocks PRs with**:
- Moderate or higher severity vulnerabilities
- GPL/AGPL/LGPL licenses
- Malicious packages

**To enable**:
Automatically runs on all pull requests to main/develop branches.

### 4. Secret Scanning
**Built-in GitHub feature**

Scans for accidentally committed secrets:
- API keys
- Tokens
- Passwords
- Private keys
- Database connection strings

**To enable** (Repository Admin only):
1. Go to Settings → Code security and analysis
2. Enable "Secret scanning"
3. Enable "Push protection" (prevents commits with secrets)

### 5. Security Policy
**Configuration**: `SECURITY.md`

Provides:
- Supported versions
- Vulnerability reporting process
- Security best practices
- Disclosure policy

### 6. Branch Protection Rules

**Recommended settings for `main` branch**:

1. **Require pull request reviews**
   - At least 1 approval required
   - Dismiss stale reviews when new commits pushed

2. **Require status checks**
   - Build must pass
   - CodeQL analysis must pass
   - Dependency review must pass
   - Unit tests must pass

3. **Require conversation resolution**
   - All review comments must be resolved

4. **Require signed commits** (optional but recommended)
   - GPG/SSH signature verification

5. **Restrict pushes**
   - Only allow admins to bypass

**To configure**:
1. Go to Settings → Branches
2. Add rule for `main` branch
3. Configure protection options

## Security Scanning Results

### Viewing Results

**CodeQL Alerts**:
- Go to Security → Code scanning alerts
- Filter by severity, status, tool
- View detailed findings with remediation guidance

**Dependabot Alerts**:
- Go to Security → Dependabot alerts
- Review vulnerable dependencies
- Auto-generate PRs to fix

**Secret Scanning Alerts**:
- Go to Security → Secret scanning alerts
- Review detected secrets
- Revoke and rotate compromised credentials

### Responding to Alerts

1. **Critical/High Severity**:
   - Review immediately
   - Create private security advisory if needed
   - Fix within 7 days
   - Release patch version

2. **Medium Severity**:
   - Review within 30 days
   - Fix in next minor release
   - Document in release notes

3. **Low Severity**:
   - Review as time permits
   - Consider for future releases
   - May be accepted as risk

## Security Workflows

### Weekly Security Review
Every Monday:
1. Dependabot creates update PRs
2. CodeQL runs full analysis
3. Review new security alerts
4. Merge approved updates

### Pull Request Security
On every PR:
1. CodeQL scans changed code
2. Dependency review checks new dependencies
3. Build must pass with zero warnings
4. Unit tests must pass
5. At least 1 review required

### Release Security
Before each release:
1. Review all open security alerts
2. Run full test suite
3. Update SECURITY.md with changes
4. Create security advisory if needed

## Best Practices

### For Contributors

1. **Never commit secrets**:
   - Use environment variables
   - Use .gitignore for sensitive files
   - Use push protection

2. **Keep dependencies updated**:
   - Review Dependabot PRs promptly
   - Test updates thoroughly
   - Merge security updates quickly

3. **Write secure code**:
   - Validate all inputs
   - Use parameterized queries
   - Avoid hard-coded credentials
   - Follow OWASP guidelines

4. **Review security alerts**:
   - Check CodeQL findings
   - Understand the vulnerability
   - Test the fix
   - Document the remediation

### For Maintainers

1. **Enable all security features**:
   - Dependabot alerts and updates
   - CodeQL scanning
   - Secret scanning with push protection
   - Dependency review

2. **Configure branch protection**:
   - Require reviews
   - Require status checks
   - Restrict pushes to main

3. **Respond to security issues**:
   - Acknowledge within 48 hours
   - Create security advisory
   - Develop fix privately
   - Coordinate disclosure

4. **Monitor security posture**:
   - Review Security tab weekly
   - Update security policy
   - Train contributors
   - Audit access regularly

## Troubleshooting

### CodeQL Configuration Conflict

**Error:** `CodeQL analyses from advanced configurations cannot be processed when the default setup is enabled`

**Cause:** The repository has GitHub's default CodeQL setup enabled, which conflicts with the custom advanced workflow.

**Solution:**
1. Go to Settings → Code security and analysis
2. Find "Code scanning" section
3. Look for "CodeQL analysis" 
4. If it shows "Default" with a gear icon:
   - Click the "..." menu
   - Select "Switch to advanced"
5. If prompted to create a workflow:
   - Don't create a new one (we already have `.github/workflows/codeql.yml`)
   - Close the dialog
6. GitHub will now use the advanced configuration from `.github/workflows/codeql.yml`
7. The custom config `.github/codeql/codeql-config.yml` will be applied automatically

**Alternative:** If you prefer default setup, delete the custom workflows:
```bash
rm .github/workflows/codeql.yml
rm -rf .github/codeql/
```
Then enable default setup through Settings → Code security and analysis.

### CodeQL False Positives

If CodeQL reports false positives:
1. Review the finding thoroughly
2. Document why it's a false positive
3. Add exclusion to `.github/codeql/codeql-config.yml`
4. Include detailed reason

Example:
```yaml
query-filters:
  - exclude:
      id: cs/specific-rule-id
      reason: "Detailed explanation of why this is safe"
```

### Dependabot PR Conflicts

If Dependabot PRs have conflicts:
1. Close the PR
2. Dependabot will recreate it
3. Or manually update the dependency

### Build Failures After Updates

If build fails after dependency update:
1. Review breaking changes in release notes
2. Update code to match new API
3. Run tests locally
4. Commit fixes to Dependabot branch

## Resources

- [GitHub Security Best Practices](https://docs.github.com/en/code-security/getting-started/github-security-features)
- [CodeQL Documentation](https://codeql.github.com/docs/)
- [Dependabot Documentation](https://docs.github.com/en/code-security/dependabot)
- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [CWE Database](https://cwe.mitre.org/)

## Support

For security-related questions:
- Security vulnerabilities: Follow SECURITY.md
- General questions: Open a discussion
- Bug reports: Open an issue

---

**Last Updated**: 2024-10-19

This security setup ensures ExcelMcp maintains high security standards and protects users from vulnerabilities.
