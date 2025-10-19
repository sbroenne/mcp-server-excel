# Security Policy

## Supported Versions

We actively support the following versions of ExcelMcp with security updates:

| Version | Supported          | Status |
| ------- | ------------------ | ------ |
| 1.0.x   | :white_check_mark: | Active |
| < 1.0   | :x:                | Unsupported |

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
- **Static Analysis**: SecurityCodeScan.VS2019 package integrated
- **CodeQL Scanning**: Automated security scanning on every push

### COM Security
- **Controlled Excel Automation**: Excel.Application runs with `Visible=false` and `DisplayAlerts=false`
- **Resource Cleanup**: Comprehensive COM object disposal and garbage collection
- **No Remote Connections**: Only local Excel automation supported

### Dependency Management
- **Dependabot**: Automated dependency updates and security patches
- **Dependency Review**: Pull request scanning for vulnerable dependencies
- **Central Package Management**: Consistent versioning across all projects

## Reporting a Vulnerability

We take security vulnerabilities seriously. If you discover a security issue, please follow these steps:

### 1. **DO NOT** Create a Public Issue
Please do not create a public GitHub issue for security vulnerabilities. This could put all users at risk.

### 2. Report Privately
Report security vulnerabilities using one of these methods:

**Preferred Method: GitHub Security Advisories**
1. Go to https://github.com/sbroenne/mcp-server-excel/security/advisories
2. Click "Report a vulnerability"
3. Fill out the advisory form with detailed information

**Alternative: Email**
Send an email to: [maintainer email - to be added]

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
PoC: ExcelMcp.exe pq-export "../../../etc/passwd" "query"
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
- **Macro Security**: VBA operations require user consent via `setup-vba-trust`

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
- **GitHub Security Advisories**: https://github.com/sbroenne/mcp-server-excel/security/advisories
- **Release Notes**: https://github.com/sbroenne/mcp-server-excel/releases
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

- **GitHub Security**: https://github.com/sbroenne/mcp-server-excel/security
- **Maintainer**: @sbroenne

## Additional Resources

- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [Microsoft Security Response Center](https://msrc.microsoft.com/)
- [CVE Database](https://cve.mitre.org/)
- [National Vulnerability Database](https://nvd.nist.gov/)

## Version History

| Version | Date | Security Changes |
|---------|------|------------------|
| 1.0.0   | 2024 | Initial security implementation with input validation |

---

**Last Updated**: 2024-10-19

Thank you for helping keep ExcelMcp and its users safe!
