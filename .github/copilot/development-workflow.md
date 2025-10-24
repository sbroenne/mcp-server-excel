# Development Workflow

> **Required process for all contributions**

## Branch Protection Rules

**⛔ NEVER commit directly to `main` branch.**

Branch protection enforces:
- ✅ PR reviews required
- ✅ CI/CD status checks must pass
- ✅ Branches must be up-to-date with main
- ❌ No force pushes or deletions

---

## Required Development Process

### 1. Create Feature Branch
```powershell
git checkout -b feature/your-feature-name
# Never work directly on main!
```

### 2. Development Standards
- **Code Quality**: Zero build warnings, all tests pass
- **Testing**: New features must include unit tests
- **Documentation**: Update README.md, COMMANDS.md, etc.
- **Security**: Follow enforced rules (CA2100, CA3003, CA3006)

### 3. PR Requirements Checklist
```
- [ ] Code builds with zero warnings (dotnet build -c Release)
- [ ] All tests pass (dotnet test)
- [ ] New features have comprehensive tests
- [ ] Documentation updated for new commands/features
- [ ] Follows existing architectural patterns
- [ ] Security rules compliance
```

### 4. Version Management
- **Don't manually update version numbers** - Release workflow handles this
- **Semantic versioning**: Major.Minor.Patch (v1.2.3)
- **Only maintainers create releases** by pushing version tags

---

## CI/CD Pipeline Strategy

### Build Workflows
- **build-cli.yml**: Builds CLI, no tests (Excel not available)
- **build-mcp-server.yml**: Builds MCP server, no tests

### Release Workflows
- **release-cli.yml**: Runs Unit tests only (`Category=Unit&RunType!=OnDemand`)
- **release-mcp-server.yml**: No tests, documents local testing

### Security Workflows
- **codeql.yml**: Security scanning, no tests
- **dependency-review.yml**: Dependency validation

---

## Test Execution Strategy

### Development
```bash
# Fast feedback (excludes OnDemand)
dotnet test --filter "Category=Unit&RunType!=OnDemand"
```

### Pre-Commit
```bash
# Comprehensive validation (excludes OnDemand)
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand"
```

### Pool Code Changes (MANDATORY)
```bash
# Verify 5 tests matched
dotnet test --filter "RunType=OnDemand" --list-tests

# Run pool cleanup tests (all must pass)
dotnet test --filter "RunType=OnDemand"
```

### CI/CD (GitHub Actions)
```bash
# Unit tests only (no Excel)
dotnet test --filter "Category=Unit&RunType!=OnDemand"
```

---

## Workflow Configuration Management

**⚠️ CRITICAL: Update ALL workflows when making config changes**

### When to Update Workflows

1. **.NET SDK Version Changes**
   - Update `global.json`
   - Update ALL workflows with `dotnet-version: X.Y.x`

2. **Assembly/Package Name Changes**
   - Update `.csproj` files
   - Update workflow executable references
   - Update package ID references

3. **Runtime Requirements**
   - Update target framework
   - Update release notes with runtime requirements

4. **Project Structure Changes**
   - Update path filters
   - Update build commands

### Validation Checklist
```powershell
# Check .NET version consistency
$globalJson = (Get-Content global.json | ConvertFrom-Json).sdk.version
$workflowVersions = Select-String -Path .github/workflows/*.yml -Pattern "dotnet-version:"

# Check assembly names match
$assemblyNames = Select-String -Path src/**/*.csproj -Pattern "<AssemblyName>"
$workflowRefs = Select-String -Path .github/workflows/*.yml -Pattern "\.exe"

# Check package IDs match
$packageIds = Select-String -Path src/**/*.csproj -Pattern "<PackageId>"
$workflowPkgRefs = Select-String -Path .github/workflows/*.yml -Pattern "tool install"
```

---

## Quality Enforcement

### Build Settings
```xml
<TreatWarningsAsErrors>true</TreatWarningsAsErrors>
<EnableNETAnalyzers>true</EnableNETAnalyzers>
<AnalysisLevel>latest</AnalysisLevel>
<EnforceCodeStyleInBuild>true</EnforceCodeStyleInBuild>
```

### Security Rules (Errors)
- **CA2100** - SQL injection prevention
- **CA3003** - File path injection prevention
- **CA3006** - Process command injection prevention
- **CA5389** - Archive path traversal prevention
- **CA5390** - Hard-coded encryption detection
- **CA5394** - Insecure randomness detection

---

## Common Mistakes to Prevent

### ❌ Don't
- Commit directly to main
- Skip writing tests
- Ignore build warnings
- Update version numbers manually
- Create releases without proper workflow

### ✅ Always
- Use feature branches
- Write comprehensive tests
- Update documentation
- Follow security best practices
- Use proper commit messages

---

## Contribution Template

```markdown
## Feature: [Brief Description]

### Changes Made
- [ ] Core implementation with COM interop
- [ ] CLI commands
- [ ] MCP server tools
- [ ] Unit + Integration + RoundTrip tests
- [ ] Documentation updates

### Testing
- [ ] All tests pass locally
- [ ] Pool cleanup tests run (if pool code changed)
- [ ] No build warnings

### Workflows Reviewed (if config changed)
- [ ] build-cli.yml
- [ ] build-mcp-server.yml
- [ ] release-cli.yml
- [ ] release-mcp-server.yml
- [ ] codeql.yml

### Breaking Changes
- None | [Description]
```

---

## Release Process (Maintainers Only)

1. Ensure all PRs merged to main
2. All CI/CD checks passing
3. Push version tag (e.g., `v1.2.3`)
4. Release workflows auto-build and publish
5. GitHub release created automatically

---

## Key Principles

1. **Feature branches mandatory** - No direct main commits
2. **Tests required** - No untested code
3. **CI/CD must pass** - Quality gates enforced
4. **Documentation updated** - Keep docs in sync
5. **Version management automated** - Don't touch manually
6. **Security enforced** - Code analysis rules as errors
