# Implementation Plan: .NET 10 Framework Upgrade

**Branch**: `006-dotnet10-upgrade` | **Date**: 2025-01-12 | **Spec**: [spec.md](spec.md)
**Input**: Feature specification from `/specs/006-dotnet10-upgrade/spec.md`

## Summary

Upgrade ExcelMcp solution from .NET 8 to .NET 10, updating SDK version, target framework across all 8 projects, CI/CD workflows, Docker images, and documentation. This is a configuration-only upgrade with no code changes required.

## Technical Context

**Language/Version**: C# 14 / .NET 10.0 (upgrading from C# 12 / .NET 8.0)  
**Primary Dependencies**: ModelContextProtocol, Microsoft.Extensions.*, Application Insights  
**Storage**: N/A (Excel files managed via COM)  
**Testing**: xUnit integration tests with real Excel instances  
**Target Platform**: Windows x64/ARM64 (COM interop requirement)
**Project Type**: Multi-project solution (4 src + 4 tests)  
**Performance Goals**: No regression from .NET 8 baseline  
**Constraints**: .NET 10 GA required (no preview), Windows-only (Excel COM)  
**Scale/Scope**: 8 csproj files, 5 workflows, 7+ documentation files, 1 Dockerfile

## Constitution Check

*GATE: Must pass before Phase 0 research. Re-check after Phase 1 design.*

| Principle | Status | Notes |
|-----------|--------|-------|
| I. Success Flag Integrity | ✅ N/A | No code changes, configuration only |
| II. MCP JSON Contract | ✅ N/A | No API changes |
| III. Tool Descriptions | ✅ N/A | No tool changes |
| IV. COM Lifecycle | ✅ N/A | No COM code changes |
| V. Exception Propagation | ✅ N/A | No exception handling changes |
| VI. COM API First | ✅ Compliant | No new dependencies |
| VII. Integration-Only Testing | ✅ Compliant | Existing tests validate upgrade |
| VIII. Test File Isolation | ✅ N/A | No test changes |
| IX. Surgical Test Execution | ✅ Compliant | Will run feature tests only |
| X. Save Only for Persistence | ✅ N/A | No test changes |
| XI. PR Workflow | ✅ Required | Must use PR for this change |
| XII. Test Before Commit | ✅ Required | Must run integration tests |
| XIII. Never Commit Automatically | ✅ Required | User approval for all commits |
| XIV. Comprehensive Bug Fixes | ✅ N/A | Not a bug fix |
| XV. Check PR Review Comments | ✅ Required | Will fix automated comments |
| XVI. Core-MCP Coverage | ✅ N/A | No new commands |
| XVII. No Placeholders | ✅ Compliant | No TODOs introduced |
| XVIII. Trust IDE Warnings | ✅ Compliant | Will address any new warnings |

**Gate Status**: ✅ PASSED - Proceed to implementation

## Project Structure

### Documentation (this feature)

```text
specs/006-dotnet10-upgrade/
├── spec.md              # Feature specification ✅
├── plan.md              # This file ✅
├── research.md          # Phase 0 output (minimal - upgrade well-documented)
├── data-model.md        # N/A - no data model changes
├── quickstart.md        # N/A - no new APIs
├── contracts/           # N/A - no API changes
└── tasks.md             # Phase 2 output (/speckit.tasks command)
```

### Source Code (repository root)

```text
# Existing structure - no changes to layout
src/
├── ExcelMcp.ComInterop/    # COM interop patterns
├── ExcelMcp.Core/          # Business logic
├── ExcelMcp.CLI/           # CLI tool
└── ExcelMcp.McpServer/     # MCP server

tests/
├── ExcelMcp.ComInterop.Tests/
├── ExcelMcp.Core.Tests/
├── ExcelMcp.CLI.Tests/
└── ExcelMcp.McpServer.Tests/
```

**Structure Decision**: No changes to project structure. This is a framework version upgrade only.

## Complexity Tracking

> No constitution violations. This is a straightforward configuration upgrade.
---

## Phase 0: Research

### R-001: Verify .NET 10 GA SDK Version

**Status**: ✅ Complete

**Finding**: .NET 10 GA released November 2024. Latest SDK version for `global.json`:
- SDK Version: `10.0.100`
- Runtime Version: `10.0.0`

**Source**: Microsoft .NET downloads page, winget package `Microsoft.DotNet.SDK.10`

### R-002: Verify NuGet Package Compatibility

**Status**: ✅ Complete

**Finding**: All current dependencies support .NET 10:
- `ModelContextProtocol` - Targets `netstandard2.0+`, compatible
- `Microsoft.Extensions.Hosting` - .NET 10 compatible packages available
- `Microsoft.ApplicationInsights.WorkerService` - .NET 10 compatible
- `xunit` / `Moq` / test packages - .NET 10 compatible

**No package version changes required** - existing versions support net10.0.

### R-003: Verify Docker Base Images

**Status**: ✅ Complete

**Finding**: .NET 10 Docker images available:
- Build: `mcr.microsoft.com/dotnet/sdk:10.0`
- Runtime: `mcr.microsoft.com/dotnet/runtime:10.0`

### R-004: Verify GitHub Actions .NET Setup

**Status**: ✅ Complete

**Finding**: `actions/setup-dotnet@v4` supports `dotnet-version: 10.0.x`

---

## Phase 1: Design

### No Design Required

This upgrade involves **configuration changes only**. No new:
- Data models
- API contracts
- User interfaces
- Business logic

### Files to Modify (Categorized)

#### Category 1: SDK & Framework (Core Changes)

| File | Current | Target | Requirement |
|------|---------|--------|-------------|
| `global.json` | `8.0.416` | `10.0.100` | FR-002 |
| `src/ExcelMcp.ComInterop/ExcelMcp.ComInterop.csproj` | `net8.0` | `net10.0` | FR-001 |
| `src/ExcelMcp.Core/ExcelMcp.Core.csproj` | `net8.0` | `net10.0` | FR-001 |
| `src/ExcelMcp.CLI/ExcelMcp.CLI.csproj` | `net8.0` | `net10.0` | FR-001 |
| `src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj` | `net8.0` | `net10.0` | FR-001 |
| `tests/ExcelMcp.ComInterop.Tests/ExcelMcp.ComInterop.Tests.csproj` | `net8.0` | `net10.0` | FR-001 |
| `tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj` | `net8.0` | `net10.0` | FR-001 |
| `tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj` | `net8.0` | `net10.0` | FR-001 |
| `tests/ExcelMcp.McpServer.Tests/ExcelMcp.McpServer.Tests.csproj` | `net8.0` | `net10.0` | FR-001 |

#### Category 2: CI/CD Workflows

| File | Change | Requirement |
|------|--------|-------------|
| `.github/workflows/build-mcp-server.yml` | `dotnet-version: 8.0.x` → `10.0.x`, paths `net8.0` → `net10.0` | FR-003, FR-004 |
| `.github/workflows/build-cli.yml` | `dotnet-version: 8.0.x` → `10.0.x`, paths `net8.0` → `net10.0` | FR-003, FR-004 |
| `.github/workflows/release-mcp-server.yml` | `dotnet-version: 8.0.x` → `10.0.x` | FR-003 |
| `.github/workflows/release-vscode-extension.yml` | `dotnet-version: 8.0.x` → `10.0.x` (if applicable) | FR-003 |
| `.github/workflows/codeql.yml` | `dotnet-version: 8.0.x` → `10.0.x` | FR-003 |

#### Category 3: Container

| File | Change | Requirement |
|------|--------|-------------|
| `Dockerfile` | Base images `sdk:8.0` → `sdk:10.0`, `runtime:8.0` → `runtime:10.0` | FR-007 |

#### Category 4: Documentation

| File | Change | Requirement |
|------|--------|-------------|
| `README.md` | Badge `.NET 8.0` → `.NET 10`, requirements section | FR-005, FR-006 |
| `docs/INSTALLATION.md` | .NET 10 requirement, winget command | FR-010, FR-011 |
| `src/ExcelMcp.McpServer/README.md` | .NET 10 requirement | FR-006 |
| `src/ExcelMcp.CLI/README.md` | .NET 10 requirement | FR-006 |
| `gh-pages/index.md` | .NET 10 requirement | FR-005, FR-006 |
| `gh-pages/installation.md` | .NET 10 requirement, winget command | FR-010, FR-011 |
| `vscode-extension/CHANGELOG.md` | Document .NET 10 requirement change | FR-015 |

#### Category 5: Already Updated

| File | Status | Notes |
|------|--------|-------|
| `.specify/memory/constitution.md` | ✅ Done | Updated to v1.1.0 with .NET 10 |

---

## Implementation Order

### Step 1: Core Framework Changes (Blocking)
1. Update `global.json` to SDK `10.0.100`
2. Update all 8 `.csproj` files to `net10.0`
3. Run `dotnet restore` to verify package compatibility
4. Run `dotnet build` to verify compilation
5. Run integration tests to verify functionality

### Step 2: CI/CD Updates (Blocking for PR)
1. Update all workflow files to `dotnet-version: 10.0.x`
2. Update artifact paths from `net8.0` to `net10.0`

### Step 3: Container Updates (Non-Blocking)
1. Update Dockerfile base images to .NET 10

### Step 4: Documentation Updates (Non-Blocking)
1. Update README.md badge and requirements
2. Update installation documentation with winget command
3. Update component READMEs
4. Update gh-pages documentation
5. Update VS Code extension CHANGELOG

### Step 5: Verification
1. Build solution with 0 warnings
2. Run integration tests (excluding VBA)
3. Verify Docker build (if applicable)
4. Create PR and verify CI/CD passes

---

## Success Verification Checklist

| Criterion | Verification Command/Action |
|-----------|----------------------------|
| SC-001: Build 0 warnings | `dotnet build --configuration Release` |
| SC-002: Tests pass | `dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA"` |
| SC-003: Workflows pass | Check GitHub Actions after PR |
| SC-004: NuGet targets net10.0 | Inspect `.nupkg` contents |
| SC-005: Docs accurate | Manual review of all updated files |
| SC-006: Docker builds | `docker build -t test .` |
| SC-007: No new warnings | Verify build output |

---

## Risk Assessment

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| Package incompatibility | Low | Medium | Spec assumption verified - packages support net10.0 |
| CI/CD path issues | Low | Medium | Comprehensive path updates in workflow files |
| New compiler warnings | Low | Low | Address any C# 14 warnings as they arise |
| Docker image availability | Low | Low | Microsoft publishes images at GA |

---

## Rollback Plan

If critical issues discovered post-merge:
1. Revert `global.json` to `8.0.416`
2. Revert all `.csproj` files to `net8.0`
3. Revert workflow `dotnet-version` to `8.0.x`
4. Revert Dockerfile base images to `8.0`

All changes are configuration-only, making rollback straightforward.
