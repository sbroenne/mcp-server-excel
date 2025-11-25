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
git tag v1.1.0

# Push the tag (triggers release workflow)
git push origin v1.1.0
```

1. **Automated Release Workflow**:
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

## 🧪 **Testing Requirements & Organization**

### **Three-Tier Test Architecture**

ExcelMcp uses a **production-ready three-tier testing approach** with organized directory structure:

```
tests/
├── ExcelMcp.Core.Tests/
│   ├── Unit/           # Fast tests, no Excel required (~2-5 sec)
│   ├── Integration/    # Medium speed, requires Excel (~1-15 min)
│   └── RoundTrip/      # Slow, comprehensive workflows (~3-10 min each)
├── ExcelMcp.McpServer.Tests/
│   ├── Unit/           # Fast tests, no server required  
│   ├── Integration/    # Medium speed, requires MCP server
│   └── RoundTrip/      # Slow, end-to-end protocol testing
└── ExcelMcp.CLI.Tests/
    ├── Unit/           # Fast tests, no Excel required
    └── Integration/    # Medium speed, requires Excel & CLI
```

### **Development Workflow Commands**

**During Development (Fast Feedback):**
```powershell
# Quick validation - run tests for specific feature
dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"
dotnet test --filter "Feature=DataModel&RunType!=OnDemand"
```

**Before Commit (Comprehensive):**
```powershell
# Full local validation - runs in 10-20 minutes (excludes VBA)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"
```

**Session/Batch Code Changes (MANDATORY):**
```powershell
# When modifying ExcelSession.cs or ExcelBatch.cs
dotnet test --filter "RunType=OnDemand"
```

### **Test Categories & Guidelines**

**⚠️ No Unit Tests** - See `docs/ADR-001-NO-UNIT-TESTS.md` for architectural rationale

**Integration Tests (`Category=Integration`)**
- ✅ Test business logic with real Excel COM interaction
- ✅ Medium speed (10-20 minutes for full suite)
- ✅ Requires Excel installation
- ✅ These ARE our unit tests (Excel COM cannot be mocked)
- ✅ Run specific features during development
- ✅ Slow execution (3-10 minutes each)
- ✅ Verifies actual Excel state changes
- ✅ Comprehensive scenario coverage

### **Adding New Tests**

When creating tests, follow these placement guidelines:

```csharp
// Unit Test Example
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class CommandLogicTests 
{
    // Tests business logic without Excel
}

// Integration Test Example  
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public class PowerQueryCommandsTests
{
    // Tests single Excel operations
}

// Round Trip Test Example
[Trait("Category", "RoundTrip")]
[Trait("Speed", "Slow")]
[Trait("Feature", "EndToEnd")]
[Trait("RequiresExcel", "true")]
public class VbaWorkflowTests
{
    // Tests complete workflows: import → run → verify → export
}
```

### **PR Testing Requirements**

Before creating a PR, ensure:

```powershell
# Required - Integration tests pass (excludes VBA)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Code builds without warnings
dotnet build -c Release

# Code follows style guidelines (automatic via EditorConfig)
```

**For Complex Features:**
- ✅ Add integration tests for all Excel operations
- ✅ Test round-trip persistence (create → save → reload → verify)
- ✅ Update documentation
- ✅ No unit tests needed (see ADR-001-NO-UNIT-TESTS.md)

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

## 📊 **Application Insights / Telemetry Setup**

ExcelMcp uses Azure Application Insights for anonymous usage telemetry and crash reporting. Telemetry is **opt-out** (enabled by default in release builds).

### **What is Tracked**

- **Tool invocations**: Tool name, action, duration (ms), success/failure
- **Unhandled exceptions**: Exception type and redacted stack trace
- **Session ID**: Random GUID per process (no user identification)

### **What is NOT Tracked**

- File paths, file names, or file contents
- User identity, machine name, or IP address
- Excel data, formulas, or cell values
- Connection strings, credentials, or passwords

### **Sensitive Data Redaction**

All telemetry passes through `SensitiveDataRedactingProcessor` which removes:
- Windows file paths (`C:\Users\...` → `[REDACTED_PATH]`)
- UNC paths (`\\server\share\...` → `[REDACTED_PATH]`)
- Connection string secrets (`Password=...` → `[REDACTED_CREDENTIAL]`)
- Email addresses → `[REDACTED_EMAIL]`

### **Azure Resources Setup (Maintainers Only)**

To deploy the Application Insights infrastructure:

```powershell
# 1. Login to Azure
az login

# 2. Deploy resources (creates RG, Log Analytics, App Insights)
.\infrastructure\azure\deploy-appinsights.ps1 -SubscriptionId "<your-subscription-id>"

# 3. Copy the connection string from output
# Output: "Connection String: InstrumentationKey=xxx;IngestionEndpoint=..."
```

### **GitHub Secret Configuration (Maintainers Only)**

After deploying Azure resources:

1. Go to GitHub repo → **Settings** → **Secrets and variables** → **Actions**
2. Add new secret: `APPINSIGHTS_CONNECTION_STRING`
3. Paste the connection string from deployment output

The release workflow automatically injects this at build time.

### **Local Development**

During local development, telemetry is **disabled by default** because the placeholder connection string is not replaced. This is intentional - no telemetry data is sent from dev builds.

#### **Debug Mode: Console Output**

To test telemetry locally without Azure, enable debug mode which logs to stderr:

```powershell
# Enable debug telemetry (logs to console instead of Azure)
$env:EXCELMCP_DEBUG_TELEMETRY = "true"

# Build and run the MCP server
dotnet build src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj
dotnet run --project src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj

# You'll see telemetry output like:
# [Telemetry] Debug mode enabled - logging to stderr
# Activity.TraceId: abc123...
# Activity.DisplayName: ToolInvocation
# Activity.Tags:
#     tool.name: excel_file
#     tool.action: list
#     tool.duration_ms: 42
#     tool.success: true
```

#### **Testing with Real Azure Resources**

To test with actual Application Insights:

```powershell
# 1. Deploy Azure resources
.\infrastructure\azure\deploy-appinsights.ps1 -SubscriptionId "<your-sub-id>"

# 2. Temporarily inject connection string (DON'T COMMIT!)
$connStr = "InstrumentationKey=xxx;IngestionEndpoint=https://..."
(Get-Content "src/ExcelMcp.McpServer/Telemetry/ExcelMcpTelemetry.cs") -replace `
    '__APPINSIGHTS_CONNECTION_STRING__', $connStr | `
    Set-Content "src/ExcelMcp.McpServer/Telemetry/ExcelMcpTelemetry.cs"

# 3. Build and run
dotnet build src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj
dotnet run --project src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj

# 4. Check Azure Portal → Application Insights → Transaction search

# 5. IMPORTANT: Revert the file (don't commit connection string!)
git checkout src/ExcelMcp.McpServer/Telemetry/ExcelMcpTelemetry.cs
```

To verify telemetry state:
```csharp
// ExcelMcpTelemetry.IsEnabled returns false when:
// - Connection string is placeholder "__APPINSIGHTS_CONNECTION_STRING__"
// - User has opted out via EXCELMCP_TELEMETRY_OPTOUT=true

// ExcelMcpTelemetry.IsEnabled returns true when:
// - EXCELMCP_DEBUG_TELEMETRY=true (console output mode)
// - Connection string is real (injected at build time)
```

### **User Opt-Out**

Users can disable telemetry by setting an environment variable:

```powershell
# Windows
$env:EXCELMCP_TELEMETRY_OPTOUT = "true"

# Or permanently via System Properties → Environment Variables
```

### **Telemetry Architecture**

```
MCP Tool Invocation
    │
    ▼
ExcelToolsBase.ExecuteToolAction()
    │ (tracks: tool, action, duration, success)
    ▼
ExcelMcpTelemetry.TrackToolInvocation()
    │
    ▼
SensitiveDataRedactingProcessor
    │ (removes: paths, credentials, emails)
    ▼
Azure Monitor Exporter → Application Insights
```

### **Files Overview**

| File | Purpose |
|------|---------|
| `Telemetry/ExcelMcpTelemetry.cs` | Static helper for tracking |
| `Telemetry/SensitiveDataRedactingProcessor.cs` | Redacts PII before transmission |
| `Program.cs` | OpenTelemetry configuration |
| `infrastructure/azure/appinsights.bicep` | Azure resource definitions |
| `infrastructure/azure/deploy-appinsights.ps1` | Deployment script |

## 📞 **Need Help?**

- **Read the docs**: [Contributing Guide](CONTRIBUTING.md)
- **Ask questions**: Create a GitHub Issue with the `question` label
- **Report bugs**: Use the bug report template

---

**Remember: Every change, no matter how small, must go through a Pull Request!**

This ensures code quality, proper testing, and maintains the project's reliability for all users.
