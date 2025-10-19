# Rename Solution from ExcelCLI to ExcelMcp

## 🎯 Objective

Rename the entire C# solution from `ExcelCLI` to `ExcelMcp` and create a new GitHub repository named `mcp-server-excel` to align with MCP ecosystem naming conventions and follow Microsoft C# naming conventions for 3+ letter acronyms (PascalCase: `Mcp` not `MCP`).

## 📋 Background

- **Current:** Repository `sbroenne/ExcelCLI` with `ExcelCLI` namespace
- **Target:** New repository `sbroenne/mcp-server-excel` with `ExcelMcp` namespace
- **Reason:** Align with MCP ecosystem naming convention (`mcp-server-{technology}`)
- **C# Convention:** Three-letter acronyms should use PascalCase (`Xml`, `Json`, `Mcp`) per [Microsoft guidelines](https://learn.microsoft.com/en-us/dotnet/standard/design-guidelines/capitalization-conventions)
- **Result:** New repo `mcp-server-excel` with solution `ExcelMcp.sln` and `ExcelMcp.*` projects

## 🔧 Required Changes

### Phase 1: Create New GitHub Repository

Create a new GitHub repository following MCP naming conventions:

**Repository Details:**

- **Name:** `mcp-server-excel`
- **Owner:** `sbroenne`
- **Description:** "Excel Development MCP Server for AI Assistants - Power Query refactoring, VBA enhancement, and Excel automation for GitHub Copilot, Claude, and ChatGPT"
- **Visibility:** Public
- **Topics/Tags:** `mcp-server`, `excel`, `power-query`, `vba`, `github-copilot`, `ai-assistant`, `excel-development`, `model-context-protocol`, `excel-automation`, `dotnet`
- **License:** MIT (copy from existing repo)
- **README:** Will be updated in Phase 2

**Actions:**

1. Create new repository `sbroenne/mcp-server-excel` on GitHub
2. DO NOT initialize with README (we'll push from existing repo)
3. Set repository description and topics as specified above
4. Configure branch protection rules for `main` branch (same as current repo)
5. Copy repository settings from `sbroenne/ExcelCLI`:
   - Issues enabled
   - Discussions enabled
   - Security alerts enabled
   - Dependabot enabled

**Note:** Analyze the current repository settings and replicate them in the new repository. The Copilot agent should identify all relevant settings that need to be transferred.

### Phase 2: Rename Solution and Projects

### 1. Solution File Rename

- **File:** `ExcelCLI.sln` → `ExcelMcp.sln`
- **Action:** Rename file and update all internal project references

### 2. Project Files & Directories

Rename these project files and their containing directories:

#### Source Projects

- `src/ExcelCLI/` → `src/ExcelMcp.CLI/`
  - `ExcelCLI.csproj` → `ExcelMcp.CLI.csproj`
- `src/ExcelCLI.Core/` → `src/ExcelMcp.Core/`
  - `ExcelCLI.Core.csproj` → `ExcelMcp.Core.csproj`
- `src/ExcelCLI.MCP.Server/` → `src/ExcelMcp.Server/`
  - `ExcelCLI.MCP.Server.csproj` → `ExcelMcp.Server.csproj`

#### Test Projects

- `tests/ExcelCLI.Tests/` → `tests/ExcelMcp.Tests/`
  - `ExcelCLI.Tests.csproj` → `ExcelMcp.Tests.csproj`

### 3. Namespace Updates

Update all C# files with namespace declarations:

**Find and replace across ALL `.cs` files:**

- `namespace ExcelCLI` → `namespace ExcelMcp.CLI`
- `namespace ExcelCLI.Commands` → `namespace ExcelMcp.CLI.Commands`
- `namespace ExcelCLI.Core` → `namespace ExcelMcp.Core`
- `namespace ExcelCLI.MCP.Server` → `namespace ExcelMcp.Server`
- `namespace ExcelCLI.Tests` → `namespace ExcelMcp.Tests`
- `using ExcelCLI` → `using ExcelMcp.CLI`
- `using ExcelCLI.Commands` → `using ExcelMcp.CLI.Commands`
- `using ExcelCLI.Core` → `using ExcelMcp.Core`
- `using ExcelCLI.MCP.Server` → `using ExcelMcp.Server`

### 4. Project Reference Updates

Update project references in all `.csproj` files:

```xml
<!-- OLD -->
<ProjectReference Include="..\ExcelCLI.Core\ExcelCLI.Core.csproj" />
<ProjectReference Include="..\ExcelCLI\ExcelCLI.csproj" />

<!-- NEW -->
<ProjectReference Include="..\ExcelMcp.Core\ExcelMcp.Core.csproj" />
<ProjectReference Include="..\ExcelMcp.CLI\ExcelMcp.CLI.csproj" />
```

### 5. Assembly Information Updates

Update assembly names in `.csproj` files:

```xml
<!-- In ExcelMcp.CLI.csproj -->
<AssemblyName>ExcelMcp</AssemblyName>
<RootNamespace>ExcelMcp.CLI</RootNamespace>

<!-- In ExcelMcp.Core.csproj -->
<AssemblyName>ExcelMcp.Core</AssemblyName>
<RootNamespace>ExcelMcp.Core</RootNamespace>

<!-- In ExcelMcp.Server.csproj -->
<AssemblyName>ExcelMcp.Server</AssemblyName>
<RootNamespace>ExcelMcp.Server</RootNamespace>
```

### 6. Documentation Updates

Update references in documentation files:

#### Files to Update

- `README.md` - Update project structure, namespace references
- `.github/copilot-instructions.md` - Update namespace examples
- `.github/workflows/build.yml` - Update project paths
- `.github/workflows/release.yml` - Update project paths and artifact names
- `docs/DEVELOPMENT.md` - Update namespace examples
- `docs/CONTRIBUTING.md` - Update project structure references
- All other `docs/*.md` files mentioning `ExcelCLI` namespaces

#### Key Updates

- Replace `ExcelCLI` namespace references with `ExcelMcp.CLI`
- Replace `ExcelCLI.Core` with `ExcelMcp.Core`
- Replace `ExcelCLI.MCP.Server` with `ExcelMcp.Server`
- Update code examples showing namespace usage
- Update project structure diagrams

### 7. CI/CD Workflow Updates

#### build.yml workflow

```yaml
# Update project paths
- name: Build ExcelMcp.CLI
  run: dotnet build src/ExcelMcp.CLI/ExcelMcp.CLI.csproj

- name: Build ExcelMcp.Core
  run: dotnet build src/ExcelMcp.Core/ExcelMcp.Core.csproj

- name: Build ExcelMcp.Server
  run: dotnet build src/ExcelMcp.Server/ExcelMcp.Server.csproj
```

#### release.yml workflow

```yaml
# Update binary paths and artifact names
- name: Publish CLI
  run: dotnet publish src/ExcelMcp.CLI/ExcelMcp.CLI.csproj

# Update file copy operations
- name: Copy binaries
  run: |
    Copy-Item "src/ExcelMcp.CLI/bin/Release/net8.0/publish/ExcelMcp.exe" ...
```

### 8. Test Project Updates

Update test namespaces and usings in all test files:

- `tests/ExcelMcp.Tests/` - All `*.cs` files
- Update `using ExcelCLI` statements to `using ExcelMcp.CLI`
- Update namespace declarations

### 9. Git Considerations

**IMPORTANT:** Use `git mv` for renames to preserve history:

```bash
# Example commands (Copilot should use these)
git mv ExcelCLI.sln ExcelMcp.sln
git mv src/ExcelCLI src/ExcelMcp.CLI
git mv src/ExcelCLI.Core src/ExcelMcp.Core
git mv src/ExcelCLI.MCP.Server src/ExcelMcp.Server
git mv tests/ExcelCLI.Tests tests/ExcelMcp.Tests
```

### 10. Repository Migration

After completing all renames and updates:

**Actions:**

1. Create feature branch for all changes: `git checkout -b feature/rename-to-excelmcp`
2. Commit all changes with comprehensive message
3. Push to new repository: `git remote add new-origin https://github.com/sbroenne/mcp-server-excel.git`
4. Push all branches: `git push new-origin --all`
5. Push all tags: `git push new-origin --tags`
6. Verify all GitHub Actions workflows execute successfully in new repository
7. Update any external references (badges, links) to point to new repository

**Note:** The Copilot agent should analyze what else needs to be done for a complete repository migration, including:

- Updating URLs in documentation
- Migrating GitHub Actions secrets (if needed)
- Updating external integrations
- Creating redirect notice in old repository (optional)
- Any other repository-specific configurations

## ✅ Acceptance Criteria

### New Repository

- [ ] New repository `sbroenne/mcp-server-excel` created
- [ ] Repository description and topics set correctly
- [ ] Branch protection rules configured
- [ ] Repository settings match original repo
- [ ] License file present (MIT)

### Build & Test
- [ ] Solution builds successfully with zero warnings
- [ ] All unit tests pass (`dotnet test`)
- [ ] All integration tests pass
- [ ] CI/CD workflows execute successfully

### Naming Consistency:
- [ ] All namespaces follow `ExcelMcp.*` pattern
- [ ] "MCP" uses PascalCase (`Mcp`) not all caps
- [ ] CLI tool explicitly named `ExcelMcp.CLI`
- [ ] No remaining `ExcelCLI` references in code

### File Structure:
- [ ] All directories renamed correctly
- [ ] All `.csproj` files renamed and updated
- [ ] Solution file references all projects correctly
- [ ] Git history preserved via `git mv`

### Documentation:
- [ ] All code examples use new namespaces
- [ ] Project structure diagrams updated
- [ ] CI/CD workflow paths updated
- [ ] README and all docs reflect new names

### Functionality:
- [ ] CLI executable works with all commands
- [ ] MCP server starts and responds correctly
- [ ] Core library exports correct assembly name
- [ ] No breaking changes to public APIs

## 🚨 Critical Notes

1. **Create new repository FIRST** before making code changes
2. **Use `git mv` for ALL renames** to preserve Git history
3. **Follow PascalCase for "Mcp"** - Not `MCP` (3+ letter acronym rule)
4. **CLI becomes `ExcelMcp.CLI`** - Explicit project naming
5. **Server becomes `ExcelMcp.Server`** - Remove redundant `.MCP.`
6. **Test after EACH major step** - Build/test incrementally
7. **Update Directory.Build.props** if it contains project-specific references
8. **Check for hardcoded paths** in any PowerShell scripts or batch files
9. **Analyze and identify** any additional changes needed beyond what's explicitly listed
10. **Complete repository migration** including all branches, tags, and settings## 📚 Reference

- [Microsoft C# Naming Guidelines](https://learn.microsoft.com/en-us/dotnet/standard/design-guidelines/capitalization-conventions)
- Current solution: `ExcelCLI.sln`
- Target solution: `ExcelMcp.sln`
- Projects: 4 total (3 source + 1 test)

## 🎯 Definition of Done

### Repository Setup

- ✅ New repository `sbroenne/mcp-server-excel` created and configured
- ✅ Repository settings, branch protection, and topics configured
- ✅ All branches and tags migrated from old repository

### Code Changes

- ✅ Solution renamed to `ExcelMcp.sln`
- ✅ All projects renamed to `ExcelMcp.*` pattern
- ✅ All namespaces updated to `ExcelMcp.*`
- ✅ All file paths updated in CI/CD workflows
- ✅ Documentation updated with new repository name and namespaces
- ✅ All URLs updated to point to new repository
- ✅ Git history preserved via `git mv`
- ✅ No references to old `ExcelCLI` namespace remain in code

### Validation

- ✅ Build succeeds with zero warnings in new repository
- ✅ All tests pass in new repository
- ✅ CI/CD workflows execute successfully in new repository
- ✅ MCP server functionality verified
- ✅ CLI functionality verified

### Additional Tasks (Agent to Identify)

- ✅ All other required changes identified and completed by Copilot agent
- ✅ Repository migration checklist verified
- ✅ No breaking changes to public APIs (where possible)

---

**Estimated Scope:** Large refactoring affecting ~100+ files + repository migration  
**Breaking Changes:** Yes - namespace changes affect external consumers  
**Git Strategy:** Feature branch with `git mv` for history preservation, then repository migration  
**Testing Required:** Comprehensive - build, unit tests, integration tests, CI/CD validation in new repository  
**Agent Autonomy:** Copilot agent should identify and complete additional required changes not explicitly listed
