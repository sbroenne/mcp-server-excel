# Feature Specification: .NET 10 Framework Upgrade

**Feature Branch**: `006-dotnet10-upgrade`  
**Created**: 2025-12-28  
**Status**: Draft  
**Input**: User description: "Upgrade the project from .NET 8 to .NET 10 target framework"

## User Scenarios & Testing *(mandatory)*

### User Story 1 - Developer Builds Project with .NET 10 SDK (Priority: P1)

As a developer, I want to build the ExcelMcp solution using .NET 10 SDK so that I can take advantage of the latest runtime improvements, language features, and long-term support.

**Why this priority**: This is the core upgrade - without successful builds, no other stories can be tested or delivered. All downstream functionality depends on the project compiling and running on .NET 10.

**Independent Test**: Clone repository, ensure .NET 10 SDK is installed, run `dotnet build` and verify all projects compile with zero warnings and zero errors.

**Acceptance Scenarios**:

1. **Given** .NET 10 SDK is installed, **When** developer runs `dotnet build` at solution root, **Then** all 8 projects (4 src + 4 tests) compile successfully with 0 warnings
2. **Given** .NET 10 SDK is installed, **When** developer runs `dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA"`, **Then** all integration tests pass
3. **Given** a fresh clone of the repository, **When** developer runs `dotnet restore`, **Then** all NuGet packages restore successfully with compatible .NET 10 versions

---

### User Story 2 - CI/CD Pipeline Builds and Tests on .NET 10 (Priority: P2)

As a maintainer, I want GitHub Actions workflows to use .NET 10 SDK so that automated builds and releases use the new target framework.

**Why this priority**: CI/CD validation is critical for quality gates and ensures all contributors use the correct SDK version.

**Independent Test**: Push a commit to trigger CI workflows and verify all workflow jobs complete successfully with .NET 10.

**Acceptance Scenarios**:

1. **Given** code is pushed to `main` branch, **When** `build-mcp-server.yml` workflow runs, **Then** it uses .NET 10 SDK and builds successfully
2. **Given** code is pushed to `main` branch, **When** `build-cli.yml` workflow runs, **Then** it uses .NET 10 SDK and builds successfully
3. **Given** a release tag is created, **When** `release-mcp-server.yml` workflow runs, **Then** NuGet packages are published with `net10.0` target framework
4. **Given** workflow artifact verification step runs, **When** checking for built executables, **Then** paths reference `net10.0` folder instead of `net8.0`

---

### User Story 3 - End Users Run MCP Server and CLI on .NET 10 Runtime (Priority: P3)

As an end user, I want the published MCP Server and CLI to run on .NET 10 runtime so that I benefit from performance improvements and modern runtime features.

**Why this priority**: This affects the distributed artifacts but depends on successful builds and CI/CD updates first.

**Independent Test**: Install published NuGet package or download release assets, run the tools, and verify they execute correctly.

**Acceptance Scenarios**:

1. **Given** user installs `Sbroenne.ExcelMcp.McpServer` NuGet package, **When** they run the server, **Then** it starts successfully and responds to MCP protocol requests
2. **Given** user installs `Sbroenne.ExcelMcp.CLI` NuGet package, **When** they run `excelcli --help`, **Then** help text displays correctly
3. **Given** README and installation docs reference .NET version, **When** user reads documentation, **Then** they see .NET 10 as the required runtime with winget installation command

---

### User Story 4 - Users Upgrading from Previous Versions (Priority: P3)

As a user upgrading from a previous version of ExcelMcp, I want clear instructions on how to install .NET 10 runtime so that I can continue using the tools without issues.

**Why this priority**: Existing users need a smooth upgrade path with clear instructions to avoid confusion.

**Independent Test**: Follow release notes instructions to install .NET 10 via winget and verify the upgraded tools work.

**Acceptance Scenarios**:

1. **Given** user has previous ExcelMcp version with .NET 8, **When** they read the release notes, **Then** they see clear instructions to install .NET 10 runtime via winget
2. **Given** user runs `winget install Microsoft.DotNet.Runtime.10`, **When** installation completes, **Then** .NET 10 runtime is available and ExcelMcp tools work correctly

---

### Edge Cases

- What happens when a developer has only .NET 8 SDK installed but not .NET 10?
  - Build should fail with a clear error message indicating .NET 10 SDK is required
- What happens when NuGet packages have .NET 10-incompatible dependencies?
  - Package restore should fail with clear dependency resolution errors
- How does the system handle the Dockerfile which uses .NET base images?
  - Dockerfile must be updated to use `mcr.microsoft.com/dotnet/runtime:10.0` or appropriate .NET 10 image

## Requirements *(mandatory)*

### Functional Requirements

- **FR-001**: All `.csproj` files MUST specify `<TargetFramework>net10.0</TargetFramework>`
- **FR-002**: `global.json` MUST specify the latest .NET 10 GA SDK version (e.g., `10.0.100`) - preview versions are NOT permitted
- **FR-003**: All GitHub Actions workflows MUST use `dotnet-version: 10.0.x` in `setup-dotnet` action
- **FR-004**: All workflow artifact paths referencing `net8.0` MUST be updated to `net10.0`
- **FR-005**: `README.md` badge MUST display ".NET 10" instead of ".NET 8.0"
- **FR-006**: `README.md` requirements section MUST state ".NET 10 runtime" requirement
- **FR-007**: `Dockerfile` MUST use .NET 10 base images for build and runtime stages
- **FR-008**: All NuGet package dependencies MUST be compatible with .NET 10 or updated to compatible versions
- **FR-009**: Constitution file (`.specify/memory/constitution.md`) MUST reflect .NET 10 SDK requirement
- **FR-010**: `docs/INSTALLATION.md` MUST reference .NET 10 runtime requirement
- **FR-011**: `docs/INSTALLATION.md` MUST include winget installation command for .NET 10 (e.g., `winget install Microsoft.DotNet.Runtime.10`)
- **FR-012**: Project MUST build with zero warnings after upgrade (preserving `TreatWarningsAsErrors=true`)
- **FR-013**: All existing integration tests MUST pass after upgrade
- **FR-014**: Release notes MUST document the .NET 10 upgrade and include winget installation instructions for users upgrading from previous versions
- **FR-015**: `vscode-extension/CHANGELOG.md` MUST document the .NET 10 runtime requirement change for the next VS Code extension release

### Files Requiring Changes

| Category | Files |
|----------|-------|
| SDK Version | `global.json` |
| Target Framework | `src/ExcelMcp.ComInterop/ExcelMcp.ComInterop.csproj`, `src/ExcelMcp.Core/ExcelMcp.Core.csproj`, `src/ExcelMcp.CLI/ExcelMcp.CLI.csproj`, `src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj`, `tests/ExcelMcp.ComInterop.Tests/ExcelMcp.ComInterop.Tests.csproj`, `tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj`, `tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj`, `tests/ExcelMcp.McpServer.Tests/ExcelMcp.McpServer.Tests.csproj` |
| CI/CD Workflows | `.github/workflows/build-mcp-server.yml`, `.github/workflows/build-cli.yml`, `.github/workflows/release-mcp-server.yml`, `.github/workflows/release-vscode-extension.yml`, `.github/workflows/codeql.yml` |
| Documentation | `README.md`, `docs/INSTALLATION.md`, `src/ExcelMcp.McpServer/README.md`, `src/ExcelMcp.CLI/README.md`, `gh-pages/index.md`, `gh-pages/installation.md`, `vscode-extension/CHANGELOG.md` (release notes) |
| Container | `Dockerfile` |
| Constitution | `.specify/memory/constitution.md` (already updated to 1.1.0) |

## Constraints

- **.NET 10 GA Required**: This upgrade MUST use the latest .NET 10 GA (General Availability) release. Preview versions are NOT permitted. ✅ .NET 10 GA is now available (released November 2025).

## Assumptions
- **Package compatibility**: All current NuGet dependencies are assumed compatible with .NET 10. If incompatibilities arise during implementation, packages will be updated to compatible versions.
- **No breaking API changes**: The upgrade is expected to be a drop-in replacement with no code changes required beyond configuration files.
- **VS Code Extension unchanged**: The VS Code extension (TypeScript-based) does not depend on .NET version and requires no changes.

## C# 14 Code Improvement Opportunities *(optional - post-upgrade enhancements)*

The .NET 10 SDK includes C# 14 with several language features that could improve the ExcelMcp codebase. These are **optional enhancements** that can be addressed in follow-up work after the core framework upgrade is complete.

### High-Value Features

| Feature | Applicability | Where to Use | Benefit |
|---------|--------------|--------------|---------|
| **`field` keyword** | ⭐⭐⭐ HIGH | `ResultBase`, `OperationResult` classes | Enforce Critical Rule #1 invariant (Success=true ⟹ ErrorMessage empty) at property level |
| **Extension members** | ⭐⭐⭐ HIGH | `ActionExtensions.cs` (12 action enums) | Cleaner extension properties instead of static methods |
| **Partial constructors** | ⭐⭐ MEDIUM | Large partial classes (`RangeCommands`, `TableCommands`, `PivotTableCommands`) | Consistent initialization across partial files |

### Detailed Analysis

#### 1. `field` Keyword for Property Invariants

**Current Pattern** (violates Critical Rule #1 if misused):
```csharp
public class ResultBase {
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
}
// No enforcement: Success=true with ErrorMessage="error" is possible
```

**With C# 14 `field` Keyword**:
```csharp
public class ResultBase {
    public bool Success { 
        get => field;
        set {
            if (value && !string.IsNullOrEmpty(ErrorMessage))
                throw new InvalidOperationException("Success cannot be true when ErrorMessage is set");
            field = value;
        }
    }
    public string? ErrorMessage { 
        get => field;
        set {
            if (Success && !string.IsNullOrEmpty(value))
                throw new InvalidOperationException("ErrorMessage cannot be set when Success is true");
            field = value;
        }
    }
}
```

**Impact**: Enforces invariant at compile-time property access, eliminating the class of bugs caught by Rule #1.

**Files affected**: `src/ExcelMcp.Core/Models/ResultTypes.cs` (40+ result classes inherit from `ResultBase`)

#### 2. Extension Members for Action Extensions

**Current Pattern**:
```csharp
public static class ActionExtensions {
    public static string ToActionString(this FileAction action) => action switch { ... };
    public static string ToActionString(this PowerQueryAction action) => action switch { ... };
    // 12 total extension methods
}
```

**With C# 14 Extension Members**:
```csharp
extension(FileAction action) {
    public string ActionString => action switch { ... };
}
extension(PowerQueryAction action) {
    public string ActionString => action switch { ... };
}
// Cleaner property-like access: action.ActionString instead of action.ToActionString()
```

**Impact**: More idiomatic property access, reduced verbosity, natural member syntax.

**Files affected**: `src/ExcelMcp.McpServer/Models/ActionExtensions.cs`

### Lower Priority Features

| Feature | Applicability | Notes |
|---------|--------------|-------|
| **Null-conditional assignment** | ⭐⭐ MEDIUM | Already using `??=` pattern in 6 places; new `?.=` could help with COM object property assignments |
| **Lambda parameter modifiers** | ⭐⭐ MEDIUM | `batch.Execute()` callbacks could use `ref`/`in` for performance-critical paths |
| **Implicit Span conversions** | ⭐ LOW | No significant string buffer operations in current codebase |
| **`nameof` with unbound generics** | ⭐ LOW | Limited use cases identified |

### Implementation Recommendation

1. **Phase 1 (This Upgrade)**: Complete the .NET 10 framework upgrade with no code changes
2. **Phase 2 (Follow-up Spec)**: Create separate feature spec for "C# 14 Code Modernization" addressing:
   - `field` keyword adoption for Result types
   - Extension members migration for Action enums
   - Partial constructor implementation

This phased approach ensures the framework upgrade is clean and testable, with code modernization as a separate, lower-risk effort.

---

## Success Criteria *(mandatory)*

### Measurable Outcomes

- **SC-001**: Solution builds successfully with 0 warnings and 0 errors on .NET 10 SDK
- **SC-002**: All integration tests pass (excluding VBA tests which require manual environment setup)
- **SC-003**: All GitHub Actions workflows pass on the upgrade PR
- **SC-004**: Published NuGet packages target `net10.0` framework
- **SC-005**: All documentation accurately reflects .NET 10 requirement
- **SC-006**: Dockerfile builds and runs successfully with .NET 10 base images
- **SC-007**: No new analyzer warnings introduced by the upgrade
