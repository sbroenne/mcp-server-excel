# Development Workflow

## 🚨 **IMPORTANT: All Changes Must Use Pull Requests**

**Direct commits to `main` are not allowed.** All changes must go through the Pull Request (PR) process to ensure:

- Code review and quality control
- Proper version management
- CI/CD validation
- Documentation updates

## 📋 **Standard Development Workflow**

### 1. **Create Feature Branch**

```powershell
# Create and switch to feature branch
git checkout -b feature/your-feature-name

# Or for bug fixes
git checkout -b fix/issue-description

# Or for documentation updates  
git checkout -b docs/update-description
```

### 2. **Make Your Changes**

```powershell
# Make code changes, add tests, update docs
# Commit frequently with clear messages

git add .
git commit -m "Add feature X with tests and documentation

- Implement core functionality
- Add comprehensive unit tests  
- Update command documentation
- Include usage examples"
```

### 3. **Push Feature Branch**

```powershell
# Push your feature branch to GitHub
git push origin feature/your-feature-name
```

### 4. **Create Pull Request**

1. Go to [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
2. Click **"New Pull Request"**
3. Select your feature branch
4. Fill out the PR template:
   - **Clear title** describing the change
   - **Detailed description** of what was changed and why
   - **Testing information** - what tests were added/run
   - **Breaking changes** - if any
   - **Documentation updates** - what docs were updated

### 5. **PR Review Process**

- **Automated checks** will run (build, tests, linting)
- **Code review** by maintainers
- **Address feedback** if requested
- **Merge** once approved and all checks pass

### 6. **After Merge**

```powershell
# Switch back to main and pull latest
git checkout main
git pull origin main

# Delete the feature branch (cleanup)
git branch -d feature/your-feature-name
git push origin --delete feature/your-feature-name
```

## 🏷️ **Release Process**

### Creating a New Release

**Only maintainers** can create releases. The process is:

1. **Ensure all changes are merged** to `main` via PRs

2. **Create and push a version tag**:

```powershell
# Create version tag (semantic versioning)
git tag v1.3.2

# Push the tag (triggers release workflow)
git push origin v1.3.2
```

**Tag Patterns:**
- MCP Server & CLI (unified): `v1.2.3`
- VS Code Extension: `vscode-v1.1.3`

3. **Automated Release Workflow**:
   - ✅ Updates version numbers in project files
   - ✅ Builds the release binaries  
   - ✅ Creates GitHub release with ZIP file
   - ✅ Updates release notes

### Version Numbering

We follow [Semantic Versioning](https://semver.org/):

- **Major** (v2.0.0): Breaking changes
- **Minor** (v1.1.0): New features, backward compatible  
- **Patch** (v1.0.1): Bug fixes, backward compatible

## 🔒 **Branch Protection Rules**

The `main` branch is protected with:

- **Require pull request reviews** - Changes must be reviewed
- **Require status checks** - CI/CD must pass
- **Require up-to-date branches** - Must be current with main
- **No direct pushes** - All changes via PR only

## 📋 **Spec-Driven Development (Spec Kit)**

### **Using Spec Kit for Feature Development**

ExcelMcp uses [GitHub Spec Kit](https://github.com/github/spec-kit) for structured feature development.

**Spec Kit Commands (in GitHub Copilot):**
```
/speckit.specify    # Generate feature specification (WHAT to build)
/speckit.plan       # Create implementation plan (HOW to build)
/speckit.tasks      # Break down into actionable tasks
/speckit.implement  # Generate code with full context
```

**Spec Structure:**
```
specs/###-feature-name/
├── spec.md    # WHAT: User stories, requirements, acceptance criteria
└── plan.md    # HOW: Architecture, tech stack, design decisions
```

**Current Specs:**
- 14 feature specifications in `specs/001-014/` directories
- Constitution in `.specify/memory/constitution.md`
- Templates in `.specify/templates/`

**Workflow:**
1. Review existing specs before starting work: `specs/001-014/`
2. Check constitution for governance rules: `.specify/memory/constitution.md`
3. Use `/speckit.specify` to create new feature specs
4. Follow spec → plan → tasks → implement workflow
5. Update specs as implementation evolves

**See:** [Spec Kit Integration Guide](.specify/README.md) for complete workflow

## 🧪 **Testing Requirements**

### **Test Architecture**

**⚠️ No Unit Tests** - See `docs/ADR-001-NO-UNIT-TESTS.md` for architectural rationale

ExcelMcp uses **integration tests as unit tests** because Excel COM cannot be meaningfully mocked.

**Test Traits:**
```csharp
[Trait("Category", "Integration")]  // All tests are integration tests
[Trait("Speed", "Medium")]           // Medium (most) or Slow (heavy operations)
[Trait("Layer", "Core")]             // Core, CLI, McpServer, or ComInterop
[Trait("Feature", "PowerQuery")]     // Feature name for targeted testing
[Trait("RequiresExcel", "true")]     // All integration tests require Excel
[Trait("RunType", "OnDemand")]       // For slow session/diagnostic tests only
```

**Valid Feature Values:**
- PowerQuery, DataModel, Tables, PivotTables, Ranges, Connections, Parameters, Worksheets, VBA, VBATrust

### **Development Workflow Commands**

**During Development (Fast Feedback):**
```powershell
# Quick validation - run tests for specific feature
dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"
dotnet test --filter "Feature=DataModel&RunType!=OnDemand"
```

**Before Commit (Comprehensive - MANDATORY per Rule 0):**
```powershell
# Full validation (10-15 minutes, excludes VBA which requires manual trust setup)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"
```

**Session/Batch Code Changes (MANDATORY):**
```powershell
# When modifying ExcelSession.cs or ExcelBatch.cs
dotnet test tests/ExcelMcp.ComInterop.Tests/ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"
```

### **Test Guidelines**

**File Isolation (CRITICAL):**
- ✅ Each test creates unique file via `CoreTestHelper.CreateUniqueTestFileAsync()`
- ❌ NEVER share test files between tests
- ✅ Use `.xlsm` for VBA tests, `.xlsx` otherwise

**Assertions:**
- ✅ Binary: `Assert.True(result.Success, $"Reason: {result.ErrorMessage}")`
- ❌ NEVER "accept both" patterns
- ✅ ALWAYS verify actual Excel state after operations

**SaveAsync Rules:**
- ❌ FORBIDDEN unless explicitly testing persistence
- ✅ ONLY for round-trip tests: Create → Save → Re-open → Verify
- ❌ NEVER call in middle of test (breaks subsequent operations)

**See:** [Testing Strategy](.github/instructions/testing-strategy.instructions.md) for complete patterns

## 📋 **MCP Server Configuration Management**

### **CRITICAL: Keep server.json in Sync**

When modifying MCP Server functionality, **you must update** `src/ExcelMcp.McpServer/.mcp/server.json`:

#### **When to Update server.json:**

- ✅ **Adding new MCP tools** - Add tool definition to `"tools"` array
- ✅ **Modifying tool parameters** - Update `inputSchema` and `properties`
- ✅ **Changing tool descriptions** - Update `description` fields
- ✅ **Adding new capabilities** - Update `"capabilities"` section
- ✅ **Changing requirements** - Update `"environment"."requirements"`

#### **server.json Synchronization Checklist:**

```powershell
# After making MCP Server code changes, verify:

# 1. Tool definitions match actual implementations
Compare-Object (Get-Content "src/ExcelMcp.McpServer/.mcp/server.json" | ConvertFrom-Json).tools (Get-ChildItem "src/ExcelMcp.McpServer/Tools/*.cs")

# 2. Build succeeds with updated configuration
dotnet build src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj

# 3. Test MCP server starts without errors
dnx Sbroenne.ExcelMcp.McpServer --yes
```

#### **server.json Structure:**

```json
{
  "version": "2.0.0",          // ← Updated by release workflow
  "tools": [                   // ← Must match Tools/*.cs implementations
    {
      "name": "excel_file",    // ← Must match [McpServerTool] attribute
      "description": "...",    // ← Keep description accurate
      "inputSchema": {         // ← Must match method parameters
        "properties": {
          "action": { ... },   // ← Must match actual actions supported
          "filePath": { ... }   // ← Must match parameter types
        }
      }
    }
  ]
}
```

#### **Common server.json Update Scenarios:**

1. **Adding New Tool:**
   ```csharp
   // In Tools/NewTool.cs
   [McpServerTool]
   public async Task<string> NewTool(string action, string parameter)
   ```
   ```json
   // Add to server.json tools array
   {
     "name": "excel_newtool",
     "description": "New functionality description",
     "inputSchema": { ... }
   }
   ```

2. **Adding Action to Existing Tool:**
   ```csharp
   // In existing tool method
   case "new-action":
     return HandleNewAction(parameter);
   ```
   ```json
   // Update inputSchema properties.action enum
   "action": {
     "enum": ["list", "create", "new-action"]  // ← Add new action
   }
   ```

## �📝 **PR Template Checklist**

When creating a PR, verify:

- [ ] **Code builds** with zero warnings
- [ ] **All tests pass** (unit tests minimum)
- [ ] **New features have tests**
- [ ] **Documentation updated** (README, etc.)
- [ ] **MCP server.json updated** (if MCP Server changes) ← **NEW**
- [ ] **Breaking changes documented**
- [ ] **Follows existing code patterns**
- [ ] **Commit messages are clear**

## 🚫 **What NOT to Do**

- ❌ **Don't commit directly to `main`**
- ❌ **Don't create releases without PRs**
- ❌ **Don't skip tests**
- ❌ **Don't ignore build warnings**
- ❌ **Don't update version numbers manually** (release workflow handles this)

## 💡 **Tips for Good PRs**

### Commit Messages

```text
✅ Good: "Add PowerQuery batch refresh command with error handling"
❌ Bad: "fix stuff"
```

### PR Titles

```text  
✅ Good: "Add batch operations for Power Query refresh"
❌ Bad: "Update code"
```

### PR Size

- **Keep PRs focused** - One feature/fix per PR
- **Break large changes** into smaller, reviewable chunks
- **Include tests and docs** in the same PR as the feature

## 🔧 **Local Development Setup**

```powershell
# Clone the repository
git clone https://github.com/sbroenne/mcp-server-excel.git
cd ExcelMcp

# Install dependencies
dotnet restore

# Run all tests
dotnet test

# Build release version
dotnet build -c Release

# Test the built executable
.\src\ExcelMcp.CLI\bin\Release\net10.0\excelcli.exe --version
```

## 📞 **Need Help?**

- **Read the docs**: [Contributing Guide](CONTRIBUTING.md)
- **Ask questions**: Create a GitHub Issue with the `question` label
- **Report bugs**: Use the bug report template

---

**Remember: Every change, no matter how small, must go through a Pull Request!**

This ensures code quality, proper testing, and maintains the project's reliability for all users.
