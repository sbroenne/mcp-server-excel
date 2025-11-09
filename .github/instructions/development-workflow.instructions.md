---
applyTo: ".github/workflows/**/*.yml,**/*.csproj,global.json"
---

# Development Workflow

> **Required process for all contributions**

## Branch Protection

**⛔ NEVER commit directly to `main`**

Enforced: PR reviews, CI/CD checks, create a branch first, up-to-date branches, no force pushes

## Development Process

1. **Create feature branch**: `git checkout -b feature/name`
2. **Standards**: Zero warnings, tests pass, docs updated, security rules followed
3. **PR Checklist**: Build passes, tests pass, docs updated, patterns followed
4. **Check PR review comments**: After creating PR, retrieve automated review feedback and fix all issues
5. **Versions**: Automated via release workflow - don't update manually

## PR Review Comment Workflow

**After creating a PR, ALWAYS check for automated review comments:**

```bash
# Retrieve inline code review comments using GitHub CLI
gh api repos/sbroenne/mcp-server-excel/pulls/PULL_NUMBER/comments --paginate

# Or use the mcp_github tool if available
mcp_github_github_pull_request_read(method="get_review_comments", owner="sbroenne", repo="mcp-server-excel", pullNumber=PULL_NUMBER)
```

**Common automated reviewers:**
- **Copilot** (code quality, performance, style)
- **github-advanced-security** (security scanning, code analysis)

**Common issues to fix:**
- Improper `/// <inheritdoc/>` on constructors/test methods that don't override
- `.AsSpan().ToString()` inefficiency - use `[..n]` range operator instead
- Nullable type access without null checks
- `foreach` → `.Select()` for functional style
- Nested if statements that can be combined
- Generic catch clauses - use specific exceptions or add justification
- Path.Combine security warnings - suppress with justification for test code

**Fix all automated review comments before requesting human review.**

## Test Execution

**See testing-strategy.instructions.md for complete test commands.**

Quick reference:
- Development: `Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust`
- Session/batch changes: `RunType=OnDemand`
- VBA tests: `(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand`

## CI/CD Workflows

**Automated on Pull Requests:**
- `build-mcp-server.yml` - Builds MCP Server on code changes
- `build-cli.yml` - Builds CLI on code changes
- `integration-tests.yml` - Runs Excel COM integration tests on Azure self-hosted runner
- `codeql.yml` - Security analysis
- `dependency-review.yml` - Dependency security scanning

**Note:** Integration tests require Excel and run on Azure VM self-hosted runner (see `docs/AZURE_SELFHOSTED_RUNNER_SETUP.md`)

## Workflow Config Updates

**⚠️ Update ALL workflows when changing:**
- .NET SDK version (`global.json` + all workflows)
- Assembly/package names (`.csproj` + workflow references)
- Runtime requirements (target framework + release notes)
- Project structure (path filters + build commands)

## Quality Enforcement

**Build Settings:** `TreatWarningsAsErrors=true`, analyzers enabled

**Security Rules (Errors):** CA2100 (SQL injection), CA3003 (file path injection), CA3006 (process injection), CA5389 (archive traversal), CA5390 (hardcoded encryption), CA5394 (insecure randomness)

## Release Process (Maintainers)

**Tag Patterns:**
- MCP Server & CLI (unified): `v1.2.3`
- VS Code Extension: `vscode-v1.1.3`

Push tag → Workflow auto-builds → GitHub release created

## Key Principles

1. Feature branches mandatory
2. Tests required
3. CI/CD must pass
4. Documentation updated
5. Version management automated
6. Security enforced
