# Tasks: .NET 10 Framework Upgrade

**Input**: Design documents from `/specs/006-dotnet10-upgrade/`
**Prerequisites**: plan.md ✅, spec.md ✅, research.md ✅

**Tests**: Not required for this feature (configuration-only upgrade, validated by existing integration tests)

**Organization**: Tasks are grouped by user story from spec.md to enable independent verification.

## Format: `[ID] [P?] [Story?] Description`

- **[P]**: Can run in parallel (different files, no dependencies)
- **[Story]**: Which user story this task belongs to (US1, US2, US3, US4)
- Exact file paths included in descriptions

---

## Phase 1: Setup

**Purpose**: Ensure prerequisites are met before making changes

- [x] T001 Verify .NET 10 SDK is installed locally by running `dotnet --list-sdks`
- [x] T002 Create feature branch `006-dotnet10-upgrade` from main (if not already on branch)

---

## Phase 2: User Story 1 - Developer Builds Project with .NET 10 SDK (Priority: P1) 🎯 MVP

**Goal**: All projects compile and tests pass with .NET 10 SDK

**Independent Test**: Run `dotnet build` and `dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA"` and verify 0 warnings, all tests pass

### Implementation for User Story 1

- [x] T003 [US1] Update SDK version in global.json from `8.0.416` to `10.0.100`
- [x] T004 [P] [US1] Update TargetFramework to net10.0 in src/ExcelMcp.ComInterop/ExcelMcp.ComInterop.csproj
- [x] T005 [P] [US1] Update TargetFramework to net10.0 in src/ExcelMcp.Core/ExcelMcp.Core.csproj
- [x] T006 [P] [US1] Update TargetFramework to net10.0 in src/ExcelMcp.CLI/ExcelMcp.CLI.csproj
- [x] T007 [P] [US1] Update TargetFramework to net10.0 in src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj
- [x] T008 [P] [US1] Update TargetFramework to net10.0 in tests/ExcelMcp.ComInterop.Tests/ExcelMcp.ComInterop.Tests.csproj
- [x] T009 [P] [US1] Update TargetFramework to net10.0 in tests/ExcelMcp.Core.Tests/ExcelMcp.Core.Tests.csproj
- [x] T010 [P] [US1] Update TargetFramework to net10.0 in tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj
- [x] T011 [P] [US1] Update TargetFramework to net10.0 in tests/ExcelMcp.McpServer.Tests/ExcelMcp.McpServer.Tests.csproj
- [x] T012 [US1] Run `dotnet restore` to verify all NuGet packages resolve correctly
- [x] T013 [US1] Run `dotnet build --configuration Release` and verify 0 warnings, 0 errors
- [x] T014 [US1] Run integration tests: `dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA"`

**Checkpoint**: User Story 1 complete - project builds and tests pass on .NET 10

---

## Phase 3: User Story 2 - CI/CD Pipeline Builds and Tests on .NET 10 (Priority: P2)

**Goal**: All GitHub Actions workflows use .NET 10 SDK and build successfully

**Independent Test**: Push commit, verify all workflow jobs complete successfully

### Implementation for User Story 2

- [x] T015 [P] [US2] Update dotnet-version from 8.0.x to 10.0.x in .github/workflows/build-mcp-server.yml
- [x] T016 [P] [US2] Update artifact paths from net8.0 to net10.0 in .github/workflows/build-mcp-server.yml
- [x] T017 [P] [US2] Update dotnet-version from 8.0.x to 10.0.x in .github/workflows/build-cli.yml
- [x] T018 [P] [US2] Update artifact paths from net8.0 to net10.0 in .github/workflows/build-cli.yml
- [x] T019 [P] [US2] Update dotnet-version from 8.0.x to 10.0.x in .github/workflows/release-mcp-server.yml
- [x] T020 [P] [US2] Update dotnet-version from 8.0.x to 10.0.x in .github/workflows/release-vscode-extension.yml (if applicable)
- [x] T021 [P] [US2] Update dotnet-version from 8.0.x to 10.0.x in .github/workflows/codeql.yml

**Checkpoint**: User Story 2 complete - CI/CD workflows configured for .NET 10

---

## Phase 4: User Story 3 - End Users Run MCP Server and CLI on .NET 10 Runtime (Priority: P3)

**Goal**: Published tools run on .NET 10 runtime, documentation is accurate

**Independent Test**: Build NuGet package, verify targets net10.0; review documentation for accuracy

### Container Update

- [x] T022 [US3] Update Dockerfile FROM statements: sdk:8.0 → sdk:10.0, runtime:8.0 → runtime:10.0

### Documentation Updates

- [x] T023 [P] [US3] Update .NET version badge from ".NET 8.0" to ".NET 10" in README.md
- [x] T024 [P] [US3] Update requirements section to state ".NET 10 runtime" in README.md
- [x] T025 [P] [US3] Update .NET version requirements in src/ExcelMcp.McpServer/README.md
- [x] T026 [P] [US3] Update .NET version requirements in src/ExcelMcp.CLI/README.md
- [x] T027 [P] [US3] Update .NET 10 requirement in gh-pages/index.md
- [x] T028 [P] [US3] Update .NET 10 requirement in gh-pages/installation.md

**Checkpoint**: User Story 3 complete - container and documentation updated

---

## Phase 5: User Story 4 - Users Upgrading from Previous Versions (Priority: P3)

**Goal**: Clear upgrade instructions with winget installation command

**Independent Test**: Follow instructions in docs to install .NET 10 via winget

### Implementation for User Story 4

- [x] T029 [US4] Update docs/INSTALLATION.md with .NET 10 requirement and winget command: `winget install Microsoft.DotNet.Runtime.10`
- [x] T030 [US4] Add winget installation instructions to gh-pages/installation.md
- [x] T031 [US4] Add entry to vscode-extension/CHANGELOG.md documenting .NET 10 runtime requirement change

**Checkpoint**: User Story 4 complete - upgrade path documented

---

## Phase 6: Verification & PR

**Purpose**: Final validation before creating pull request

- [x] T032 Run `dotnet build --configuration Release` and confirm 0 warnings, 0 errors
- [x] T033 Run integration tests: `dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA"`
- [x] T034 Verify Docker build (optional): `docker build -t excelmcp-test .`
- [x] T035 Review all modified files for accuracy and completeness
- [ ] T036 Create pull request with comprehensive description
- [ ] T037 Check and fix any automated PR review comments (Copilot, GitHub Security)
- [ ] T038 Verify all GitHub Actions workflows pass on PR

---

## Dependencies & Execution Order

### Phase Dependencies

- **Setup (Phase 1)**: No dependencies - verify prerequisites first
- **User Story 1 (Phase 2)**: Depends on Setup - CORE UPGRADE (MVP)
- **User Story 2 (Phase 3)**: Depends on User Story 1 (workflows need correct target framework)
- **User Story 3 (Phase 4)**: Can start after User Story 1 (documentation can parallel User Story 2)
- **User Story 4 (Phase 5)**: Can start after User Story 1 (documentation can parallel others)
- **Verification (Phase 6)**: Depends on all user stories complete

### Within Each User Story

- T003 (global.json) must complete before T004-T011 (csproj files)
- T004-T011 (csproj files) can all run in parallel
- T012-T014 (verification) must follow T003-T011
- Documentation tasks (T023-T031) can all run in parallel

### Parallel Opportunities

```text
# After T003 (global.json), all csproj updates can run together:
T004, T005, T006, T007, T008, T009, T010, T011

# All workflow updates can run in parallel:
T015, T016, T017, T018, T019, T020, T021

# All documentation updates can run in parallel:
T023, T024, T025, T026, T027, T028, T029, T030, T031
```

---

## Implementation Strategy

### MVP First (User Story 1 Only)

1. Complete Setup (T001-T002)
2. Complete User Story 1 (T003-T014)
3. **STOP and VALIDATE**: Build + tests pass locally
4. Can deploy/demo .NET 10 builds immediately

### Full Delivery

1. User Story 1: Core framework upgrade (blocking)
2. User Story 2: CI/CD updates (parallel with US3/US4 documentation)
3. User Story 3 + 4: Documentation updates (can run in parallel)
4. Verification: Final checks and PR

### Total Task Count

| Phase | Tasks | Parallelizable |
|-------|-------|----------------|
| Setup | 2 | 0 |
| User Story 1 (P1) | 12 | 8 |
| User Story 2 (P2) | 7 | 7 |
| User Story 3 (P3) | 7 | 6 |
| User Story 4 (P3) | 3 | 0 |
| Verification | 7 | 0 |
| **Total** | **38** | **21** |

---

## Notes

- All [P] tasks can run in parallel (different files, no dependencies)
- No new tests required - existing integration tests validate the upgrade
- Rollback plan documented in plan.md if critical issues discovered
- Constitution already updated to v1.1.0 with .NET 10 SDK requirement
