# Quickstart: Upgrade MCP SDK to 0.5.0-preview.1

**Date**: 2025-12-13  
**Spec**: `specs/001-upgrade-mcp-sdk/spec.md`

---

## Prerequisites

- Windows 10+ with Excel desktop installed (COM interop tests require Excel).
- .NET SDK 8.0 or later (`dotnet --version`).
- Git CLI.
- PowerShell 7+ (for pre-commit scripts).

---

## Step 1: Ensure Clean Baseline

```powershell
git status  # Should be on branch 001-upgrade-mcp-sdk with clean working tree
dotnet restore
dotnet build --no-restore
```

All three commands must succeed with **0 warnings** before proceeding.

---

## Step 2: Bump Dependency

Update `Directory.Packages.props` (or relevant package version file):

```diff
- <PackageVersion Include="ModelContextProtocol" Version="0.4.*" />
+ <PackageVersion Include="ModelContextProtocol" Version="0.5.0-preview.1" />
```

Then restore and build:

```powershell
dotnet restore
dotnet build --no-restore
```

Record any compiler errors/warnings as the authoritative breaking-change list.

---

## Step 3: Fix Compiler Breaks

Address each CS* error identified in Step 2:

| Error | Fix |
|-------|-----|
| Removed factory/interface | Replace with updated API per changelog |
| Obsolete enum schema (MCP9001) | Migrate to new type per SDK guidance |
| RequestOptions parameter | Pass `RequestOptions` bag instead of positional params |
| Signature changes | Update method calls (e.g., `SetLoggingLevelAsync`, `UnsubscribeRequestParams`) |
| `cancellationToken:` rename | Change named argument to `cancellationToken` |

Repeat build until success with **0 warnings**.

---

## Step 4: Run Feature-Scoped Tests

```powershell
# MCP Server tests (fast, no Excel)
dotnet test tests/ExcelMcp.McpServer.Tests/ExcelMcp.McpServer.Tests.csproj

# CLI tests (minimal Excel)
dotnet test tests/ExcelMcp.CLI.Tests/ExcelMcp.CLI.Tests.csproj

# Feature filter example (replace with actual feature as needed)
dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"
```

All tests must pass.

---

## Step 5: MCP Server Smoke Check

Run MCP server locally:

```powershell
dotnet run --project src/ExcelMcp.McpServer -- --help
# or run via stdio transport
$env:MCP_LOG_LEVEL="Debug"
dotnet run --project src/ExcelMcp.McpServer
```

Confirm:
- No output to stdout (stderr only for logs).
- Exit code `0` on normal shutdown; `1` on fatal.

---

## Step 6: Commit & Push

```powershell
git add -A
git commit -m "Upgrade ModelContextProtocol to 0.5.0-preview.1"
git push origin 001-upgrade-mcp-sdk
```

Open PR; review automated checks.

---

## Rollback

If blockers emerge after merge:

```powershell
git revert HEAD --no-edit
git push origin 001-upgrade-mcp-sdk
```

Alternatively, restore old package version and open follow-up issue.
