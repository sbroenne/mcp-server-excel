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
4. **Versions**: Automated via release workflow - don't update manually

## Test Execution

```bash
# Development (fast - excludes VBA tests)
dotnet test --filter "Category=Unit&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Pre-commit (comprehensive - excludes VBA tests)
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Session/batch code changes (MANDATORY)
dotnet test --filter "RunType=OnDemand"

# VBA tests (manual only - requires VBA trust enabled)
dotnet test --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"

# CI/CD (no Excel, no VBA)
dotnet test --filter "Category=Unit&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"
```

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
- MCP Server: `v1.2.3`
- CLI: `cli-v1.2.3`
- VS Code Extension: `vscode-v1.1.3`

Push tag → Workflow auto-builds → GitHub release created

## Key Principles

1. Feature branches mandatory
2. Tests required
3. CI/CD must pass
4. Documentation updated
5. Version management automated
6. Security enforced
