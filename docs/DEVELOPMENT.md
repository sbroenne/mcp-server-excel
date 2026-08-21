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
- Add focused integration tests
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
- **Merge strategy**: **Squash merge** — All PRs are merged as a single commit to `main` to keep history clean
  - When you click "Merge pull request," select the squash merge option
  - Verify the final commit message accurately describes your changes
  - This ensures `main` has a clean, linear history with one commit per feature/fix

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

2. Run **Release All Components** from GitHub Actions and select a semantic
   version bump or custom version.
3. The workflow compiles pending changesets, updates versions, builds and
   publishes all deliverables, creates the tag, and creates the GitHub release.

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

### **Integration-First Test Architecture**

ExcelMcp tests Excel behavior through real COM automation. Unit tests that mock
Excel do not validate the threading, type conversion, persistence, or resource
management failures that matter in production.

```
tests/
├── ExcelMcp.Core.Tests/
│   └── Integration/    # Feature and round-trip tests against real Excel
├── ExcelMcp.ComInterop.Tests/
│   └── Integration/    # Session, batch, timeout, and shutdown behavior
├── ExcelMcp.Diagnostics.Tests/
│   └── Integration/Diagnostics/ # Research tests, manual only (excluded from CI)
├── ExcelMcp.McpServer.Tests/
│   └── Integration/    # Protocol and generated-surface behavior
└── ExcelMcp.CLI.Tests/
    └── Integration/    # CLI and daemon behavior
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

When creating tests, use real Excel and all required traits:

```csharp
// Integration test example
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public class PowerQueryCommandsTests
{
    // Opens a real workbook and verifies observable Excel behavior.
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

Before the CLI project replaces its output, local builds call
`scripts\Stop-ExcelMcpProcesses.ps1` once. The script delegates cleanup to
`excelcli service stop`. Cleanup is scoped to `EXCELMCP_CLI_PIPE` and validates
tracked PID start times. Excel identities remain recorded for their daemon
generation through shutdown, so a failed final Excel exit cannot lose ownership
metadata. Daemon, startup, and tracker mutexes use separate semantic namespaces
over a case-insensitive SHA-256 pipe identity, so case variants, suffixes,
separators, special characters, and long pipe names cannot collide across
pipes or mutex roles. If the normal
CLI binary predates ownership-lifecycle sources, the
script first builds a current cleanup client in an isolated temporary output;
this stops the tracked daemon before its loaded normal-output assemblies are
replaced. A true first build with no CLI binary remains a safe no-op. Cleanup
never sweeps all Excel processes. `Test-E2E.ps1` and
`Test-CliWorkflow.ps1` allocate a private pipe for each invocation so parallel
worktrees cannot stop each other's daemon or Excel instances.

**For Complex Features:**
- ✅ Add integration tests for all Excel operations
- ✅ Test round-trip persistence (create → save → reload → verify)
- ✅ Update documentation
- ✅ No unit tests needed (see ADR-001-NO-UNIT-TESTS.md)

## 🔧 **CLI Command Code Generation**

### **Architecture Overview**

The CLI uses **Roslyn source generators** to automatically generate command classes from Core's service definitions, ensuring 1:1 parity with MCP tools:

```
Core Generator (`ServiceRegistryGenerator`)
  ↓
  Generates ServiceRegistry.{Category} classes
  Generates RouteFromSettings() bridge method
  ↓
CLI Generator (`CliSettingsGenerator`)
  ↓
  Discovers referenced [ServiceCategory] interfaces and their metadata
  Generates one command class per category (inheriting ServiceCommandBase<T>)
  Generates CliCommandRegistration.RegisterCommands()
  ↓
Program.cs calls CliCommandRegistration.RegisterCommands(config)
```

### **How It Works**

**1. Core Generator Output** (`ServiceRegistry.{Category}.g.cs`):
- Nested class `CliSettings` with all [Argument] properties
- Method `RouteFromSettings()` that maps CliSettings → service command
- Constants: `CliCommandName`, `ValidActions`, `RequiresSession`

**2. CLI Generator** (`CliSettingsGenerator.cs`):
- Discovers `[ServiceCategory]` interfaces and their emitted metadata in referenced Core assemblies (no hard-coded command list).
- For each category, generates a command class:
  ```csharp
  internal sealed class SheetCommand : ServiceCommandBase<ServiceRegistry.Sheet.CliSettings>
  {
      protected override string? GetSessionId(Settings s) => s.SessionId;
      protected override string? GetAction(Settings s) => s.Action;
      protected override IReadOnlyList<string> ValidActions => ServiceRegistry.Sheet.ValidActions;
      protected override (string, object?) Route(Settings s, string action) 
          => ServiceRegistry.Sheet.RouteFromSettings(action, s);
  }
  ```
- Generates `CliCommandRegistration.RegisterCommands()`:
  ```csharp
  public static void RegisterCommands(IConfigurator config)
  {
      config.AddCommand<SheetCommand>("worksheet")
          .WithDescription(...);
      // ... generated commands for every discovered category
  }
  ```

### **Adding a New Command Category**

When adding a new service category to Core:

1. **Add `[ServiceCategory]` interface** in Core
2. **Rebuild** - generators automatically produce:
   - ServiceRegistry class in Core
   - Command class in CLI.Generated
   - Registration entry in CliCommandRegistration
3. **Test** - verify `excelcli COMMAND_NAME --help` works.

**For Complex Features:**
- ✅ Add integration tests for all Excel operations
- ✅ Test round-trip persistence (create → save → reload → verify)
- ✅ Update documentation
- ✅ No unit tests needed (see ADR-001-NO-UNIT-TESTS.md)


## 📋 **MCP Registry Manifest**

`src/ExcelMcp.McpServer/.mcp/server.json` describes the published NuGet package
for the MCP Registry. Tool schemas are generated from Core interfaces and are
not duplicated in this manifest.

Update `server.json` only when package identity, transport, repository metadata,
or runtime requirements change. The release workflow updates its version fields.

```powershell
# Build succeeds with the packaged manifest
dotnet build src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj

# Server starts without errors
dotnet run --project src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj
```

## 📝 **PR Template Checklist**

When creating a PR, verify:

- [ ] **Code builds** with zero warnings
- [ ] **Relevant integration tests pass**
- [ ] **New features have tests**
- [ ] **Documentation updated** (README, etc.)
- [ ] **MCP server.json updated** (only if package manifest metadata changes)
- [ ] **Changeset added** for user-visible changes, or `skip-changelog` applies
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

ExcelMcp uses Azure Application Insights (Classic SDK with WorkerService integration) for anonymous usage telemetry and crash reporting. Telemetry is **opt-out** (enabled by default in release builds).

### **How It Works**

The Application Insights connection string is **embedded at build time** via MSBuild - there is no runtime environment variable lookup.

**Build-time flow:**
1. MSBuild reads `AppInsightsConnectionString` property (from `Directory.Build.props.user` or env var)
2. Generates `TelemetryConfig.g.cs` with the connection string as a `const string`
3. Compiled assembly contains the embedded connection string

### **What is Tracked**

- **Tool invocations**: Tool name, action, duration (ms), success/failure
- **Unhandled exceptions**: Exception type and redacted stack trace
- **User ID**: SHA256 hash of machine identity (anonymous, 16 chars)
- **Session ID**: Random GUID per process (8 chars)

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

### **Local Development with Telemetry**

To enable telemetry in local builds:

```powershell
# 1. Copy the template file
Copy-Item "Directory.Build.props.user.template" "Directory.Build.props.user"

# 2. Edit Directory.Build.props.user and add your connection string
# <AppInsightsConnectionString>InstrumentationKey=xxx;IngestionEndpoint=...</AppInsightsConnectionString>

# 3. Build - connection string is embedded at compile time
dotnet build src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj

# 4. Run - telemetry is automatically sent to Azure
dotnet run --project src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj
```

**Note:** `Directory.Build.props.user` is gitignored - your connection string won't be committed.

### **Local Development without Telemetry**

If you don't create `Directory.Build.props.user`, builds will have an empty connection string and telemetry will be disabled. This is the default for local development.

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

The release workflow sets this as an environment variable, and MSBuild embeds it at build time.

### **Telemetry Architecture**

```text
Build Time:
  MSBuild → reads AppInsightsConnectionString → generates TelemetryConfig.g.cs

Runtime:
  MCP Tool Invocation
      │
      ▼
  ExcelMcpTelemetry.TrackToolInvocation()
      │ (tracks: tool, action, duration, success)
      ▼
  SensitiveDataRedactingProcessor
      │ (removes: paths, credentials, emails)
      ▼
  TelemetryClient → Application Insights
```

### **Files Overview**

| File | Purpose |
|------|---------|
| `Telemetry/ExcelMcpTelemetry.cs` | Static helper for tracking events |
| `Telemetry/ExcelMcpTelemetryInitializer.cs` | Sets User.Id and Session.Id on telemetry |
| `Telemetry/SensitiveDataRedactingProcessor.cs` | Redacts PII before transmission |
| `Program.cs` | Application Insights WorkerService configuration |
| `ExcelMcp.McpServer.csproj` | MSBuild target that generates TelemetryConfig.g.cs |
| `Directory.Build.props.user.template` | Template for local dev connection string |
| `infrastructure/azure/appinsights.bicep` | Azure resource definitions |
| `infrastructure/azure/deploy-appinsights.ps1` | Deployment script |

## ✂️ **Trimming and Native AOT Compatibility**

### **Why Trimming Is Not Supported**

ExcelMcp **cannot be trimmed** due to fundamental architectural constraints of Excel COM automation. The IL trimmer removes unused code at publish time, but Excel COM interop requires dynamic code paths that the trimmer cannot statically analyze.

### **Technical Constraints**

**1. Runtime COM Activation**
```csharp
// This code CANNOT be trimmed - Excel type comes from Windows Registry at runtime
Type? excelType = Type.GetTypeFromProgID("Excel.Application");
dynamic excel = Activator.CreateInstance(excelType)!;
```

The trimmer cannot know:
- What types will be returned by `GetTypeFromProgID` (it's a Windows Registry lookup)
- What members will be called on the `dynamic` object

**2. Late-Bound COM Calls**
```csharp
// All Excel operations use dynamic dispatch - the trimmer can't trace these calls
dynamic workbook = excel.Workbooks.Open(filePath);
dynamic sheet = workbook.Worksheets.Item(1);
sheet.Range["A1"].Value2 = "Hello";
```

**3. Excel is External**
- Excel is not a .NET assembly - it's an out-of-process COM server
- The .NET runtime uses the Dynamic Language Runtime (DLR) for all Excel calls
- No static type information exists for the trimmer to analyze

### **What We DID Modernize**

While the Excel automation core cannot be trimmed, we modernized the OLE Message Filter to use .NET source-generated COM interop:

| Component | Before | After |
|-----------|--------|-------|
| `IOleMessageFilter` | `[ComImport]` | `[GeneratedComInterface]` |
| `OleMessageFilter` | `class` | `[GeneratedComClass]` partial class |
| `CoRegisterMessageFilter` | `[DllImport]` | `[LibraryImport]` |

**Benefits:**
- ✅ Compile-time marshalling code generation
- ✅ No runtime IL stub generation for the message filter
- ✅ Better diagnostics and debugging

### **Suppressed Warnings**

The following warnings are suppressed in `Directory.Build.props` because they cannot be fixed:

| Warning | Reason |
|---------|--------|
| `IL2026` | Reflection/dynamic code incompatible with trimming |
| `IL3050` | Code incompatible with Native AOT |
| `CA1416` | Windows-only APIs (this is a Windows-only project) |

### **Can We Ever Support Trimming?**

**No**, unless one of these happens:
1. **Excel exposes a .NET API** - Microsoft would need to create a managed Excel SDK
2. **We abandon COM** - Would require a completely different architecture (file-based only, no live automation)
3. **Excel is replaced** - Use a different spreadsheet engine with .NET bindings

**The current architecture is the standard approach** for Excel automation in .NET and is used by thousands of applications. Trimming is simply not compatible with COM automation.

### **Alternatives for Smaller Binaries**

If deployment size is a concern:
- Use **framework-dependent** deployment (default) - smallest option (~15 MB)
- The .NET runtime is typically already installed on Windows machines with Excel
- Self-contained deployment is only needed for isolated environments

## 📞 **Need Help?**

- **Read the docs**: [Contributing Guide](CONTRIBUTING.md)
- **Ask questions**: Create a GitHub Issue with the `question` label
- **Report bugs**: Use the bug report template

---

**Remember: Every change, no matter how small, must go through a Pull Request!**

This ensures code quality, proper testing, and maintains the project's reliability for all users.
