# Security Policy

## Supported Versions

ExcelMcp ships frequent releases (multiple per month). We only support the **latest published version** with security fixes — there are no parallel maintenance branches for older minor/patch releases:

| Version              | Supported          |
| --------------------- | ------------------ |
| Latest release         | :white_check_mark: |
| Any older release      | :x: Please upgrade  |

Check the [Releases page](https://github.com/sbroenne/mcp-server-excel/releases) for the current latest version, and keep your installation up to date via the CLI's built-in auto-update, `npx skills`, the VS Code extension, or NuGet.

## Security Features

ExcelMcp includes several security measures:

### Input Validation

- **Path Traversal Protection**: All file paths are validated with `Path.GetFullPath()`
- **File Size Limits**: 1GB maximum file size to prevent DoS attacks
- **Extension Validation**: Only `.xlsx` and `.xlsm` files are accepted
- **Path Length Validation**: Maximum 32,767 characters (Windows limit)

### Code Analysis

- **Enhanced Security Rules**: CA2100, CA3003, CA3006, CA5389, CA5390, CA5394 enforced as errors
- **Treat Warnings as Errors**: All code quality issues must be resolved
- **CodeQL Scanning**: C#, JavaScript/TypeScript, Python, and GitHub Actions scanning on pull requests to `main`, pushes to `main`, merge groups, and a weekly schedule

### COM Security

- **Controlled Excel Automation**: Excel.Application runs with `Visible=false` and `DisplayAlerts=false`
- **Resource Cleanup**: Comprehensive COM object disposal and garbage collection
- **No Remote Connections**: Only local Excel automation supported

### ExcelMcp Service Security

The ExcelMcp Service manages Excel COM automation sessions:

**MCP Server**: The service runs fully **in-process** — no inter-process communication. There is no attack surface beyond the MCP Server process itself.

**CLI**: The CLI daemon uses a **Windows named pipe** (`excelmcp-cli-{USER_SID}`) for communication between CLI commands and the daemon process:

| Protection | Status | Description |
|------------|--------|-------------|
| **User Isolation** | ✅ Enforced | Pipe name includes user SID. Users cannot access each other's daemon. |
| **Windows ACLs** | ✅ Enforced | Named pipe restricts access to current user's SID via `PipeSecurity` ACLs. |
| **Local Only** | ✅ Enforced | Named pipes are local IPC only - no network access possible. |
| **Process Restriction** | ❌ Not Enforced | Any process running as the same user can connect to the CLI daemon. |

**What This Means:**

1. **Same-user access**: Any application running under your Windows user account can connect to the CLI daemon and execute Excel operations. This is by design, similar to how Docker and database servers work.

2. **No cross-user access**: User A cannot connect to User B's CLI daemon. Each user has a separate named pipe with their SID.

3. **No network access**: The named pipe is strictly local. Remote processes cannot connect.

**Security Implications:**

- If malware runs under your user account, it could theoretically connect to the CLI daemon and control Excel
- However, such malware could already control Excel directly (or do anything else you can do)
- The service does not elevate privileges or provide capabilities beyond what the user already has

### Dependency Management

- **Dependabot**: Automated dependency updates and security patches
- **Dependency Review**: Pull request scanning for vulnerable dependencies
- **Central Package Management**: Consistent versioning across all projects

### Code scanning configuration

The [CodeQL workflow](https://github.com/sbroenne/mcp-server-excel/blob/main/.github/workflows/codeql.yml) uses **advanced setup** with the default queries plus `security-extended`, configured once in [codeql-config.yml](https://github.com/sbroenne/mcp-server-excel/blob/main/.github/codeql/codeql-config.yml).

- C# uses a manual Release solution build on Windows with the SDK from `global.json`. This includes source-generated Service, CLI, and MCP code. Excel is not needed for scanning; the workflow does not run COM tests.
- JavaScript/TypeScript, Python, and Actions use separate no-build scans on Linux, covering the extension, tooling, website code, and workflows.
- There are no workflow path filters, so documentation-only pull requests also receive security checks. The weekly scan detects newly supported issues even when code has not changed.
- Actions are pinned to full commit SHAs and kept current by the existing GitHub Actions Dependabot updates. Only analysis jobs receive `security-events: write`; checkout does not persist credentials. Contributor code runs under `pull_request`, never `pull_request_target`.
- Results upload directly to GitHub code scanning, without a second downloadable SARIF artifact. Review findings in the repository's Security tab.

The shared configuration preserves six C# query exemptions established to control false-positive noise:

| Exempt query | Reason |
| --- | --- |
| `cs/catch-of-all-exceptions` | Intentional MCP/CLI error boundaries and COM cleanup |
| `cs/invalid-dynamic-call` | Late-bound Excel COM calls |
| `cs/call-to-gc` | Existing COM resource cleanup |
| `cs/useless-cast-to-self` | Previously noisy compiler-generated regex code |
| `cs/useless-assignment-to-local` | Previously noisy compiler-generated regex code |
| `cs/complex-block` | Previously noisy compiler-generated regex code |

These are query-wide C# exemptions, not file-scoped suppressions: CodeQL query filters select query metadata, and a nested `paths` key does not restrict an exemption to source locations. They do not suppress queries for other languages. Keep these exemptions unless a review of actual findings justifies changing them; a successful analysis run alone does not establish that false positives are resolved.

C# manual-build coverage is determined by the build, not `paths`/`paths-ignore`. The solution still includes tests and generated code in extraction; the old path exclusions did not exclude them from a manual C# build. Keeping that coverage allows other security queries to inspect generated entry points. For additional false positives, prefer individual dismissals with a rationale over broader query exclusions.

### Required repository settings (administrator checklist)

Workflow files **do not enforce repository settings**. An administrator must verify and maintain the following separately:

1. **Settings → Code security → Code scanning:** use **advanced setup**, with default setup disabled. Do not enable both. After changing languages, confirm successful analyses for all four categories: `/language:csharp`, `/language:javascript-typescript`, `/language:python`, and `/language:actions`, on both `main` and a pull request.
2. **Settings → Rules → Rulesets → main:** keep the ruleset active and retain existing CI requirements. Add **Require code scanning results → CodeQL**, with **Security alerts: High or higher** and **Alerts: Errors** as the starting policy. This blocks qualifying new findings and missing/in-progress analysis; it is not a gate on every historical alert.
3. Require the four **Analyze Code with CodeQL (...)** job status checks and **dependency-review**, selecting GitHub Actions as the expected source, after successful runs make them available. This also blocks workflow execution failures. Keep `CI Gate` and `Docs Site` required. If enabling a merge queue, ensure every required workflow supports `merge_group`; the CodeQL workflow already does, but code scanning rulesets themselves do not apply to merge queue groups.
4. Set the **Code scanning pull request check** failure threshold consistently with the ruleset. A successful analysis/upload job does **not** mean no vulnerabilities were found. `security-policy`, `minimum-severity`, `fail-on-severity`, and license checks are not supported CodeQL configuration controls; dependency/license enforcement belongs in Dependency Review.
5. **Settings → Actions → General:** use read-only default workflow permissions, keep permission to create/approve pull requests disabled unless specifically needed, and require approval for outside-contributor workflows. Do not send secrets or write tokens to untrusted fork runs. Before enforcing SHA pinning repository-wide, migrate the remaining workflows; pinning the CodeQL workflow alone is not sufficient.
6. **Settings → Code security:** verify the dependency graph, Dependabot alerts and security updates, secret scanning and push protection, and private vulnerability reporting are enabled where available. Dependabot's version-update YAML does not prove these repository features are enabled.
7. Review the tool status and open alerts regularly, including scheduled-run failures and justified dismissals. After merging scanning changes, establish a successful `main` baseline and recheck a pull request before treating merge protection as verified.

Verification requires repository administration/security access. An API `403` means settings or alerts could not be inspected; it does not mean scanning is disabled or that there are no alerts.

See GitHub's [workflow configuration reference](https://docs.github.com/en/code-security/reference/code-scanning/workflow-configuration-options), [merge protection guidance](https://docs.github.com/en/code-security/how-tos/find-and-fix-code-vulnerabilities/manage-your-configuration/set-merge-protection), and [Actions security guidance](https://docs.github.com/en/actions/reference/security/secure-use).

## Reporting a Vulnerability

We take security vulnerabilities seriously. If you discover a security issue, please follow these steps:

### 1. **DO NOT** Create a Public Issue

Please do not create a public GitHub issue for security vulnerabilities. This could put all users at risk.

### 2. Report Privately

Report security vulnerabilities using one of these methods:

**Preferred Method: GitHub Security Advisories**

1. Go to <https://github.com/sbroenne/mcp-server-excel/security/advisories>
2. Click "Report a vulnerability"
3. Fill out the advisory form with detailed information

**Alternative: GitHub Direct Message**

Contact the maintainer via GitHub: [@sbroenne](https://github.com/sbroenne)

Subject: `[SECURITY] ExcelMcp Vulnerability Report`

### 3. Information to Include

Please provide as much information as possible:

- **Description**: Clear description of the vulnerability
- **Impact**: What could an attacker do with this vulnerability?
- **Affected Versions**: Which versions are affected?
- **Proof of Concept**: Steps to reproduce (if possible)
- **Suggested Fix**: If you have a fix or mitigation (optional)

Example:

```
Vulnerability: Path traversal in file operations
Impact: Attacker could read/write files outside intended directory
Affected Versions: 1.0.0 - 1.0.2
PoC: excelcli powerquery view --file "../../../etc/passwd" --query-name "Sales"
Suggested Fix: Validate resolved paths are within allowed directories
```

### 4. What to Expect

- **Acknowledgment**: Within 48 hours
- **Initial Assessment**: Within 5 business days
- **Status Updates**: Regular updates on progress
- **Fix Timeline**:
  - Critical: 7 days
  - High: 30 days
  - Medium: 90 days
  - Low: Best effort

### 5. Coordinated Disclosure

We follow responsible disclosure practices:

1. **Private Fix**: We'll develop a fix privately
2. **Security Advisory**: Create GitHub Security Advisory
3. **CVE Assignment**: Request CVE if applicable
4. **Public Release**: Release patch with security notes
5. **Credit**: We'll credit you in the release notes (if desired)

## Security Best Practices for Users

### MCP Server Security

- **Validate AI Requests**: Review Excel operations requested by AI assistants
- **File Path Restrictions**: Only allow MCP Server access to specific directories
- **Audit Logs**: Monitor MCP Server operations in logs
- **Trust Configuration**: Only enable VBA trust when necessary

### CLI Security

- **Script Validation**: Review automation scripts before execution
- **File Permissions**: Ensure Excel files have appropriate permissions
- **Isolated Environment**: Run in sandboxed environment when processing untrusted files
- **Excel Security Settings**: Maintain appropriate Excel macro security settings

### Development Security

- **Code Review**: All changes require review before merge
- **Branch Protection**: Main branch protected with required checks
- **Signed Commits**: Consider using signed commits (recommended)
- **Least Privilege**: Run with minimal required permissions

## Known Security Considerations

### Excel COM Automation

- **Local Only**: ExcelMcp only supports local Excel automation
- **Windows Only**: Requires Windows with Excel installed
- **Excel Process**: Creates Excel.Application COM objects
- **Macro Security**: VBA operations require the user to manually enable "Trust access to the VBA project object model" in Excel Trust Center settings

### File System Access

- **Full Path Resolution**: All paths resolved to absolute paths
- **No Network Paths**: UNC paths and network drives not supported
- **Current User Context**: Operations run with current user permissions

### AI Integration (MCP Server)

- **Trusted AI Assistants**: Only use with trusted AI platforms
- **Request Validation**: Review operations before Excel executes them
- **Sensitive Data**: Avoid exposing workbooks with sensitive data to AI assistants
- **Audit Trail**: MCP Server logs all operations

## Security Updates

Security updates are published through:

- **GitHub Security Advisories**: <https://github.com/sbroenne/mcp-server-excel/security/advisories>
- **Release Notes**: <https://github.com/sbroenne/mcp-server-excel/releases>
- **NuGet Advisories**: Package vulnerabilities shown in NuGet

Subscribe to repository notifications to receive security alerts.

## Vulnerability Disclosure Policy

### Our Commitment

- We will acknowledge receipt of vulnerability reports within 48 hours
- We will keep reporters informed of progress
- We will credit researchers in security advisories (if desired)
- We will not take legal action against researchers following responsible disclosure

### Researcher Guidelines

- **Responsible Disclosure**: Give us time to fix before public disclosure
- **No Harm**: Do not access, modify, or delete other users' data
- **Good Faith**: Act in good faith to help improve security
- **Legal**: Follow all applicable laws

## Security Contacts

- **GitHub Security**: <https://github.com/sbroenne/mcp-server-excel/security>
- **Maintainer**: @sbroenne

## Additional Resources

- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [Microsoft Security Response Center](https://msrc.microsoft.com/)
- [CVE Database](https://cve.mitre.org/)
- [National Vulnerability Database](https://nvd.nist.gov/)

## Version History

| Version | Date | Security Changes |
|---------|------|------------------|
| 1.7.0   | 2026 | Named pipe security with Windows ACL user isolation |
| 1.0.0   | 2025 | Initial security implementation with input validation |

---

**Last Updated**: 2026-07-09

Thank you for helping keep ExcelMcp and its users safe!
