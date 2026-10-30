# Development Workflow

## Build and check ownership

Use PowerShell 7.4+ and the SDK selected by `global.json`. Start with the
[setup guide](agents/development.md) and install the normal
[commit hook](PRE-COMMIT-SETUP.md). The working loop is editor/compiler feedback,
selected Release build, quick commit checks, CI, and independent review.

For affected checks:

```powershell
& .\build.ps1 plan --base-ref origin/main --output artifacts\validation-plan.json
& .\build.ps1 validate --plan artifacts\validation-plan.json
```

Inspect the plan first. The validation command runs the selected full local
plan, not the reduced commit checks. CI owns supported process, tooling, npm
and distribution checks; local process runs are appropriate when their boundary
changes, not as a blanket pre-push requirement. Run selected real-Excel checks
locally because ordinary hosted runners do not have desktop Excel. Commands
and explicit test groups are in [tests/README](../tests/README.md).

`scripts/Get-ValidationPlan.ps1` selects checks through explicit changed-path
and feature-label rules, not source-reference analysis or rebuild decisions.
`build.ps1` calls shared PowerShell coordination directly; there is no custom
C# build executable, SDK bootstrap or tool-output cache. MSBuild evaluates
projects/imports and owns restore, dependency builds and incremental compilation.
Selection and CI completion need only PowerShell, so their Actions jobs do not
install the .NET SDK.

The commit hook builds selected inputs and uses the `Commit` projection of
`Fast`: CLI, MCP Server, Core and Service, excluding `AdapterTestKind=System`
and `Speed=Slow`. CI retains the full `Fast`, `Process` and `Tooling`
partitions. A shared input can require a complete solution build even though the
hook executes only deterministic product tests. In CI, the first selected group
in Fast/Process/Tooling order owns that complete build; other groups build their
selected projects and dependencies. `Build-CiInputs.ps1` performs those builds.

The MCP compilation checks documentation counts through EXL030 using the
compiled Core manifest. There is no separate source scanner or count updater.
The existing xUnit architecture test exports the built Excel inventory.
Standard `dotnet test` uses `tests/strict.runsettings` to reject empty selections,
skips and runner warnings. Shared PowerShell adds bounded process ownership and
checks discovered-versus-executed names for Excel and acceptance runs.
Test-only edits select their owning project group; unit and recording-only
protocol tests do not select Excel. Real-Excel MCP transport fixtures and their
shared helpers select Excel coverage. Shared runtime helpers intentionally select broader
runtime and acceptance coverage rather than maintaining a second source analyser.
Package preparation belongs to the existing PowerShell commands.
PR CI runs npm checks and prepares only distributions affected by package inputs.
Main/manual CI and release preparation remain complete. Runtime-source-only PRs
retain product checks but defer distribution-only defects to complete main checks.
Explicit `Build-ReleasePackages.ps1 -Components ...` runs may narrow components.
The same packaging commands are available for local troubleshooting; a complete
local package run is not required for every edit or commit. Actions runs cheap
skill/extension source checks before expensive packaging and retains checks of
the finished distributions. Publication stays in authorized Actions workflows.

For the constraints behind this workflow, including completion checks, output
isolation and owned-process cleanup, follow [build rules](../scripts/AGENTS.md).

Release/publication procedures: [release strategy](RELEASE-STRATEGY.md),
[extension development](../vscode-extension/DEVELOPMENT.md), and
[plugin publication](../.github/workflows/docs/publish-plugins-setup.md#maintenance-and-updates).

## Investigating a slow workbook close

Close responses automatically include `closeTimingsMs`: a compact object mapping
each measured step to its duration in milliseconds, rounded to two decimal
places. No environment setting or local timing files are needed.
`excelcli session close` returns this object under `data.closeTimingsMs`;
the MCP `file` tool's `close` action returns `closeTimingsMs` directly.

Measurements cover readiness, saving, workbook close, release of Excel objects,
Excel thread exit, and `worker-completion`: the final work in the background
copy of ExcelMcp after the workbook request finishes. Only steps
that ran are included: for example, `save` is absent when saving was not
requested. Repeated work within a step is accumulated. Some steps include other
steps, so do not add all durations together.

The timing object is normally smaller than 1 KB. It contains only step names
and durations, not workbook data, paths, session IDs, or exception messages.
It is returned to the calling agent, not sent as usage reporting.
Failures retain their original error and any collected measurements; a duration
does not establish that its step succeeded or that pending Excel work stopped.
Closing returns after any requested save, confirmed workbook close, applicable
Excel Quit, and release of the acquired Excel objects. It does not force general
memory cleanup or wait for the Excel process to disappear. After confirmed
release, that Excel process is no longer tracked. Failures and uncertain work
retain their recovery information.

An already-empty background copy of ExcelMcp stops through the existing
stop-after-reply mechanism. Its completion reply carries the final status, so
normal close needs no separate status or stop request. Its existing record
keeps that completion available until the controller reads it, even if the
background copy exits first. This does not start a new process or exit watcher.
Saving rules and protection of other workbooks are unchanged.

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

2. Run **[Release] All Components** from GitHub Actions and select a semantic
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
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj -c Release --filter 'RequiresExcel=true&Feature=PowerQuery&RunType!=OnDemand' --blame-hang-timeout 5m --logger trx
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj -c Release --filter 'RequiresExcel=true&Feature=DataModel&RunType!=OnDemand' --blame-hang-timeout 5m --logger trx
```

**Before Commit:** Rerun affected tests and applicable repository checks.
The hook selects staged changes with the same PowerShell policy used by PR CI, including
affected dependencies. It does not automatically run every workbook feature or
all acceptance scenarios after a feature-local edit. Full acceptance remains
available explicitly and is selected for shared runtime boundaries.
Use the [changed-area commands](../tests/README.md#changed-path-ci-selection);
they retain hard deadlines and reject empty or failed selected test reports.

**Session/Batch Code Changes (MANDATORY):**
```powershell
# When modifying ExcelSession.cs or ExcelBatch.cs
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"
```

### **Test Categories & Guidelines**

Use real Excel integration tests for COM behavior and focused non-COM tests for
pure parsing, mapping, serialization, and generation.
[ADR-001](ADR-001-TESTING-STRATEGY.md) explains the choice;
the [testing instructions](../tests/AGENTS.md) define how to apply it.

**Integration Tests (`Category=Integration`)**
- ✅ Test business logic with real Excel COM interaction
- ✅ Run time depends on the selected cases and local Excel installation
- ✅ Requires Excel installation
- ✅ Mocks do not establish Excel COM behavior
- ✅ Run specific features during development
- ✅ Verifies actual Excel state changes
- ✅ Comprehensive scenario coverage

### **Adding New Tests**

When creating COM integration tests, use real Excel and all required traits.
For pure logic, follow nearby non-COM tests and mark `RequiresExcel=false`.

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
# Example: select the project and feature affected by the change
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj -c Release --filter 'RequiresExcel=true&Feature=PowerQuery&RunType!=OnDemand' --blame-hang-timeout 5m --logger trx

# Code builds without warnings
dotnet build -c Release

# Code follows style guidelines (automatic via EditorConfig)
```

Before the CLI project replaces its output, local builds call
`scripts\Stop-ExcelCliService.ps1` once. It forcibly stops background CLI services
whose executable is this worktree's Debug or Release output, after checking the
`service run` command and PID/start-time identity. It retains a native process
handle through verification and termination, so a reused PID cannot redirect
termination to another process. **Builds do not save workbooks;
unsaved work in those development sessions may be lost.** No temporary helper is
built, and no Excel, MCP, foreground CLI, or other worktree's process is stopped.
No matching service is a no-op; a real query or termination failure fails the
build. Hosted CI skips the step; self-hosted desktops retain it.

The optional `-PipeName` argument limits test-run cleanup to the selected pipe.
Builds stop all eligible local CLI services, regardless of `EXCELMCP_CLI_PIPE`.
`Test-E2E.ps1` and `Test-CliWorkflow.ps1` allocate a private pipe per invocation.
Normal `excelcli service stop` and CLI/MCP workbook save/close safeguards remain
unchanged; the direct force-stop policy applies only to development tooling.

**For Complex Features:**
- ✅ Add integration tests for all Excel operations
- ✅ Test round-trip persistence when saving is the behavior under test
- ✅ Update documentation
- ✅ Add non-COM tests for Excel-independent behavior

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
- ✅ Test round-trip persistence when saving is the behavior under test
- ✅ Update documentation
- ✅ Add non-COM tests for Excel-independent behavior


## MCP tool implementation

Feature MCP tools come from the Core contract and source generator. This includes
worksheet actions with and without an open session, and screenshot image results.
Only file/session tools retain handwritten routing. MCP calls its own controller
through the Service bridge. The controller routes Excel work to private workers
hosting the shared Service; MCP does not connect to the CLI daemon.

The official SDK registers tools with `WithToolsFromAssembly` and supplies
dependencies and cancellation tokens. Generated tools use the shared async
execution boundary and generated Service routes. Session requirements come from
each Core action's batch parameter; a tool containing both kinds of action has
an optional session input, checked for each requested action. An explicit
`ServiceCategory.CommandName` preserves an existing CLI/batch name when it differs
from the MCP tool name.

The generated screenshot tool selects the MCP image converter from its declared
`ScreenshotResult` return type. The converter packages image content and metadata
only; it owns no tool definition, input checks, or Excel access.

`ExecuteToolActionAsync` supplies common telemetry and result conversion. Use shared
`JsonOptions` for additional payloads. Tool failures set the actual SDK
`CallToolResult.IsError` and return structured JSON as well as JSON text; preserve
Service categories, HRESULTs, inner context, and retry information. Unexpected
exceptions and cancellation propagate to the SDK. Request filters reject unknown,
misspelled, and action-inapplicable arguments. An inapplicable argument whose
value is null or equals its advertised schema default is ignored, because clients
may send every advertised default. The filters also reject omitted fields listed as required in the tool's
input schema with a structured `InvalidInput` response naming the missing field;
the SDK still owns binding and injected parameters.

Each host owns its bridge through dependency injection. Ordinary shutdown closes
saved owned workbooks and leaves unsaved ones open and visible without saving or
discarding. Attached workbooks stay untouched. Explicit `file close` defaults to `save:false` and
discards edits. A cancelled or expired open/create retains its original execution,
eventual workbook, session identity, and path claim without resetting unrelated
sessions. The affected Excel window is revealed when possible. Cancellation is
not a transaction, undo, or permission to close the workbook.

Protocol cancellation tests send `notifications/cancelled` explicitly and verify
the server's retained ownership and eventual outcome. Cancelling an SDK client's local wait alone does not prove
that it sent this notification: the 2.2.0 client was observed to stop waiting
without sending it. Keep this client-side limitation separate from server cancellation handling
and do not replace the SDK's server cancellation handling with a custom protocol.

Tool and parameter descriptions should explain server-specific behavior,
constraints, and differences between overlapping tools. Types and enum values
already appear in the schema. Keep destructive/read-only metadata accurate.
Core XML documentation is passed to the MCP generator as an MSBuild additional
file so compiled interface references retain their parameter descriptions.
Known top-level parameter names are rendered in MCP snake_case; nested JSON
keys and enum values keep their contract spelling. Do not delete Core XML
documentation before downstream builds. Published MCP output still excludes it.
All MCP stdio diagnostics, including bootstrap/startup output, go to stderr;
stdout is reserved for JSON-RPC.

The server exposes tools, not prompts or resources, and sends no MCP elicitation
requests. Consent advice in descriptions and instructions must be handled by
the client; it is not an enforced server dialog. For generated
skills and minimal server instructions, see
[Maintaining skills and server guidance](../docs/AGENT-SKILLS.md#authoring-and-packaging).

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

# Build release version
dotnet build -c Release

# Run integration tests for the affected feature (requires Excel).
dotnet test tests\ExcelMcp.Service.Tests\ExcelMcp.Service.Tests.csproj -c Release --filter 'RequiresExcel=true&Feature=PowerQuery&RunType!=OnDemand' --blame-hang-timeout 5m --logger trx

# Test the built executable
.\src\ExcelMcp.CLI\bin\Release\net10.0\excelcli.exe --version
```

## 📊 **Application Insights / Telemetry Setup**

ExcelMcp uses Azure Application Insights for pseudonymous usage telemetry. Both entry
points use OpenTelemetry trace and log providers with Microsoft's direct Azure
Monitor exporter. The MCP Server also reports sanitized failure classifications. Telemetry is **opt-out** (enabled by
default in release builds with an embedded connection string). Contributor builds
default to off. Both entry points support the same `EXCELMCP_TELEMETRY_MODE`
environment variable:

| Value | Destination |
|---|---|
| Unset | Azure if the build contains a connection string; otherwise off |
| `off` | No telemetry collection or delivery |
| `local` | Approved operation events and request traces to a loopback OTLP receiver, never Azure |
| `azure` | Existing Azure delivery; requires an embedded connection string |

An explicit `off` or `local` overrides even a release package's embedded Azure
configuration. Invalid settings are reported on stderr and return a nonzero exit
code. `local` constructs only OTLP exporters, with no Azure exporter or automatic collectors.

### **How It Works**

The Azure connection string is **embedded at build time** via MSBuild. Runtime
destination selection does not replace that string or require Azure credentials.

**Build-time flow:**
1. MSBuild reads `AppInsightsConnectionString` property (from `Directory.Build.props.user` or env var)
2. Generates `TelemetryConfig.g.cs` with the connection string as a `const string`
3. Compiled assembly contains the embedded connection string

### **What is Tracked**

- **Tool operations**: Tool name, action, CLI or MCP Server entry point,
  duration (ms), success/failure. CLI records describe daemon execution, not
  client-side parsing, connection, or output handling. The daemon records each
  received operation once; internal readiness pings are excluded.
- **Unhandled exceptions (MCP Server only)**: Exception type, approved source, and project-owned
  failure site; messages and stack traces are not transmitted
- **User ID**: Random ID stored at
  `%LOCALAPPDATA%\ExcelMcp\telemetry-user-id`, shared between the CLI and MCP
  Server. Background setup reads/creates the file only when telemetry is enabled;
  disabled reporting never accesses it. Delete it and restart to reset the ID. If storage is unavailable, the
  current process uses a temporary random ID.
- **Migration link**: If telemetry is active when a new ID is first stored,
  one event sent to the active destination carries the prior name-derived hash
  so existing usage can be mapped to the new ID. Raw machine and user names are
  not sent. Disabled reporting does not create an ID; local migration links are
  never deferred to Azure. Like all reports, the link can be lost at exit.
- **Session ID**: Random GUID per process (8 chars)
### **What is NOT Tracked**

- File paths, file names, or file contents
- Raw user identity, machine name, or IP address
- Excel data, formulas, or cell values
- Connection strings, credentials, or passwords
- Excel workbook `workbook_session_id`/`sessionId` values or raw MCP arguments. The
  standard random MCP server process telemetry session ID is separate from
  workbook identity.

### **Sensitive Data Protection**

Tool telemetry contains only allowlisted operation metadata. Explicit exception
telemetry is rebuilt from safe classifications and never contains the original
message or stack. Sanitized failures are `UnhandledException` custom events.
An ingestion-time workspace transform rejects raw exception rows, framework
traces, metrics/counters, and unmarked `UnhandledException` events. `SensitiveDataRedactor` is retained for
local diagnostic text and recognizes:
- Windows file paths (`C:\Users\...` → `[REDACTED_PATH]`)
- UNC paths (`\\server\share\...` → `[REDACTED_PATH]`)
- Connection string secrets (`Password=...` → `[REDACTED_CREDENTIAL]`)
- Email addresses → `[REDACTED_EMAIL]`

### **Local Development with Telemetry**

Use the standalone Aspire dashboard as the first receiver for local telemetry
development, following the [telemetry validation rules](../AGENTS.md#build-and-validation).
No Azure resource, subscription access, or Aspire AppHost integration is needed.
Start the dashboard with its UI and OTLP listener bound to loopback, and build
Release without a production connection string. Configure the application with
the dashboard's actual OTLP port and protocol; the example below assumes gRPC
on port 4317. This pipeline does not send OTLP authentication headers, so use a
loopback-only listener configured to accept local development traffic.

The [standalone Aspire dashboard](https://aspire.dev/dashboard/standalone/)
can run without Docker or an AppHost:

```powershell
npx -y @microsoft/aspire-cli dashboard run --frontend-url http://127.0.0.1:18888 --otlp-grpc-url http://127.0.0.1:4317 --otlp-http-url http://127.0.0.1:4318
```

Open the login URL printed by Aspire to inspect telemetry. Keep that terminal
running for the local check and stop it when finished.

In the terminal that launches ExcelMcp:

```powershell
$env:EXCELMCP_TELEMETRY_MODE = 'local'
$env:OTEL_EXPORTER_OTLP_PROTOCOL = 'grpc'
$env:OTEL_EXPORTER_OTLP_ENDPOINT = 'http://localhost:4317'
.\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe session list --quiet
```

CLI telemetry is owned by its existing background service. Use a private pipe
via `EXCELMCP_CLI_PIPE`, and start that service with `excelcli service start` from
the same local-mode environment before running `session list`. An already-running
service retains the environment it inherited at startup; a stopped-service check
does not start reporting. Stop only the private test service with
`excelcli service stop` when finished.

For MCP, set these same variables on the process launched by your MCP client,
using `src\ExcelMcp.McpServer\bin\Release\net10.0-windows\Sbroenne.ExcelMcp.McpServer.exe`
as the server command. Call `file_read` with action `list` to exercise telemetry
without opening a workbook. No console exporter is used: MCP standard output
remains exclusively JSON-RPC.

The supported protocols are `grpc` (default endpoint `http://localhost:4317`)
and `http/protobuf` (default endpoint `http://localhost:4318`). The general
`OTEL_EXPORTER_OTLP_ENDPOINT` is a base URL; HTTP export appends `/v1/logs`
and `/v1/traces`. Only `localhost` or a loopback IP is accepted, without URL
credentials, queries, or fragments. Per-signal endpoint overrides and OTLP
headers are not used by this local pipeline.
Local delivery also bypasses HTTP proxies and does not follow redirects.

Events appear as structured logs and requests as traces. Only the already
approved event properties and sanitized exception classifications are recorded;
framework logs and automatic HTTP/runtime metrics are not enabled. This is not a
replica of Azure's custom-event tables, Users/Sessions views, or ingestion rules.

Exercise both entry points using synthetic inputs and inspect the received logs
and traces in Aspire. Check a normal short-lived CLI invocation without adding
an artificial delay, and check an MCP tool call during a running server session.
Confirm operation names, timings, outcomes, and approved properties; report any
missing records rather than treating a successful command as successful telemetry
delivery. Separately test with the receiver stopped or stalled to establish that
startup, commands, and shutdown do not wait for network traffic. Automated tests
below supplement this local check, not replace it. Aspire cannot validate Azure's
ingestion rules or analytics queries; record those as separate validation.

Reporting setup (including saved-ID storage) runs in one background task. Calls
skip reporting until it is ready. Only standard SDK batch processors are used;
full buffers drop records. Exit signals stop acceptance without flushing,
disposing providers, or waiting for setup/network. CLI records are queued in the
existing daemon and can be sent after the short-lived client exits. Records
pending when the daemon or MCP Server itself exits can still be lost. CLI
commands that never reach the daemon do not produce telemetry. Delivery failure never changes command results
or emits raw SDK diagnostics. Providers are separate from the MCP host lifetime.
Sampling (including rate-limited sampling), offline disk storage, live metrics,
SDK self-statistics, and automatic standard/performance metrics are disabled.

Automated tests capture telemetry in memory and test HTTP/protobuf delivery
and gRPC against isolated loopback receivers, including blocked setup/export and
full buffers. They require no cloud access:

```powershell
dotnet test tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj -c Release --filter 'Feature=Telemetry'
dotnet test tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj -c Release --filter 'Feature=Telemetry'
```

Native Azure envelope mapping is tested with the real exporter against loopback.
Workspace transformations, cloud queries, and portal reports still require
maintainer deployment verification; local tests do not contact Azure.

### **Local Development without Telemetry**

Ordinary contributor builds without an embedded connection string disable
telemetry by default. To disable it explicitly, including in released packages:

```powershell
$env:EXCELMCP_TELEMETRY_MODE = 'off'
```

Remove the variable to restore the build's default behavior.

### **Azure Resources Setup (Maintainers Only)**

To deploy the Application Insights infrastructure:

```powershell
# 1. Login to Azure
az login

# 2. Deploy resources (creates RG, Log Analytics, App Insights)
.\infrastructure\azure\deploy-appinsights.ps1 -SubscriptionId "<your-subscription-id>"

# 3. Use the connection details saved in the ignored appinsights.secrets.local file.
# Do not print, commit, or publish that file.
```

### **GitHub Secret Configuration (Maintainers Only)**

After deploying Azure resources:

1. Go to GitHub repo → **Settings** → **Secrets and variables** → **Actions**
2. Add new secret: `APPINSIGHTS_CONNECTION_STRING`
3. Set the connection string from the ignored local deployment output file

The release workflow sets this as an environment variable, and MSBuild embeds it at build time.

### **Public Usage Analytics Automation**

The weekly `usage-analytics.yml` workflow queries aggregate telemetry through a
read-only Azure workload identity, gives GitHub Copilot only the sanitized JSON,
validates every generated numeric claim, and persists the report to the
`analytics-data` branch. The Pages build restores that report read-only.

One-time maintainer setup:

```powershell
.\infrastructure\azure\configure-analytics-oidc.ps1

# Fine-grained personal token with Account permission: Copilot Requests (read)
gh secret set COPILOT_GITHUB_TOKEN
```

Use a manual dry run before enabling publication:

```powershell
gh workflow run usage-analytics.yml -f publish=false
```

After inspecting the artifact, publish a validated run with `publish=true`.

#### Effort levels and share of work

The report shows each feature's **share of actions** (every action counts once)
next to its **share of work** (each action multiplied by a fixed effort level).
`.github/usage-analytics-weights.json` is the single source of truth for those
levels, for which homepage feature each tool belongs to, and for how CLI command
categories map to MCP tools. The collection script reads it to build its
queries, so no level or mapping is copied into the scripts.

Pick a level by the kind of work Excel does, not by measured duration (duration
mostly reflects workbook size and the user's machine):

| Level | Weight | Use for |
| --- | --- | --- |
| `light` | 1 | Reads, lists, and lookups; window actions; CLI service and diagnostic commands |
| `medium` | 3 | Changes to cells, formats, sheets, tables, charts, names, or code; screenshots and image export |
| `heavy` | 10 | Power Query, Data Model, and connection refresh or evaluation; creating queries, models, relationships, and PivotTables; running macros; what-if analysis |

Actions that existed only in older releases have no level. The report lists
them as `unweightedActions` and leaves them out of share of work. Opening and
closing workbooks remain excluded through `excludedActions`.

The report also compares command line (`excelcli`) and AI assistant (MCP
Server) use, starting from the release that first recorded the entry point.
A group with fewer than `entryPointMinimumUsers` users shows only that it lacks
enough data.

The `habits` section of the report describes how people work: session size and
areas used together, whether new users return, typical wait by area, weekday
versus weekend use, and how soon people first try advanced areas. Session
figures cover the AI assistant (MCP Server) only, because every `excelcli`
command runs as its own process. Waits depend on workbook size and the
user's machine, so they show which areas are heavier, not product speed.
Every habit group needs at least `minimumUsers` users, and small groups are
handled in one of two ways. Session size, weekday, and returning-user groups
that are too small are published only as `enoughData: false` but still count
toward totals and averages. Areas used together, typical wait by area, and
first use of advanced areas leave out too-small rows entirely. First use of
advanced areas covers only people first seen in the last
`firstAdvancedUseWindowDays` days, well inside the 90-day log retention.

Reliability (failures and errors) is not part of this report; it is covered by
a separate reliability report.

The `downloads` section reads public download counters without extra secrets:
npm (daily history for the `@sbroenne/mcp-server-excel` and
`@sbroenne/excelcli` launchers; the platform packages they install are left out
to avoid double counting), GitHub release files (`.zip`,
`.vsix`, and `.mcpb` only; checksums and release metadata are fetched by
automation), and VS Code Marketplace installs. GitHub and the
Marketplace only publish running totals, so the collect job restores the last
published report and `Update-UsageAnalytics.ps1 -PreviousReportPath` carries
its dated `snapshots` forward, adds today's totals (replacing a same-day entry),
and stores the gain between snapshots in `weeklyGains`, with `days` between the
two snapshots. When reports are not six to eight days apart, the page charts the
average week of that gap and marks the bar. Downloads count
downloads, not people, and the channels overlap. npm weekly history covers the
last 52 full weeks.

NuGet is deliberately left out. Every version, even old ones nobody would pick,
collects roughly the same couple of hundred downloads from automated mirrors and
scanners, so the totals mostly reflect how many versions were released. The
npm launchers, standalone executables, VS Code extension, and Claude Desktop
bundle never download from NuGet, and the version check asks GitHub.

The written summary is generated by GitHub Copilot CLI with the model pinned by
`Invoke-UsageAnalyticsReport.ps1 -Model` (default `claude-opus-5.5`), so weekly
summaries do not change when the CLI default model changes.

Test the scripts locally without Azure access:

```powershell
.\scripts\Test-UsageAnalytics.ps1
```

Azure custom events remain in `AppEvents` and operation traces in `AppRequests`.
Safe identity fields (`UserId`, `SessionId`, `AppVersion`) are explicit
`Properties` attributes; analytics queries normalize `Properties.UserId` with a
fallback to the historical `UserId` column before applying `UserIdMigration`.
Native non-HTTP operation traces have response code `0`; use `Success` and
`Properties.Outcome` for outcome analytics rather than HTTP response codes.
Reports count received operations, not a guaranteed complete or sampled population.

### **Telemetry Architecture**

```text
Build Time:
  MSBuild → reads AppInsightsConnectionString → generates TelemetryConfig.g.cs

Runtime:
  CLI daemon request or MCP Tool Invocation
      │
      ▼
  CliTelemetry / ExcelMcpTelemetry
      │ (tracks: tool, action, entry point, duration, success)
      ▼
  Allowlisted telemetry construction
      │ (exception messages and stacks are omitted)
      ▼
  UsageTelemetry → SDK batch processors → direct Azure Monitor exporter
      → ingestion privacy transform → Application Insights
  Or: explicit local mode → OTLP logs/traces → local dashboard
  Or: off → no collection or delivery
```

### **Files Overview**

| File | Purpose |
|------|---------|
| `Telemetry/ExcelMcpTelemetry.cs` | Static helper for tracking events |
| `ExcelMcp.CLI/Telemetry/CliTelemetry.cs` | CLI daemon operation telemetry and lifecycle |
| `Shared/UsageTelemetry.cs` | Background setup, no-wait recording/exit, native Azure or local OTLP batch export |
| `Shared/UsageRecord.cs` | Safe event and operation fields shared by both adapters |
| `Shared/TelemetrySettings.cs` | Destination selection and loopback validation |
| `Shared/TelemetryUserIdentity.cs` | Shared persisted random identity and one-time migration |
| `Telemetry/SensitiveDataRedactor.cs` | Redacts sensitive local diagnostic text |
| Entry-point `Program.cs` files | Telemetry destination and lifecycle configuration |
| CLI and MCP Server project files | MSBuild targets that generate `TelemetryConfig.g.cs` |
| `Directory.Build.props.user.template` | Template for local dev connection string |
| `infrastructure/azure/appinsights-resources.bicep` | Azure resources and ingestion privacy transforms |
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
