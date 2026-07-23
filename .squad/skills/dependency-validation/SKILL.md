# Dependency Validation

## When to use

Use this skill when validating dependency or toolchain upgrades in `mcp-server-excel`, especially when the touched surface includes `Directory.Packages.props`, `global.json`, `vscode-extension\package.json`, or release packaging paths.

## Goal

Get a trustworthy answer fast: did the upgrade break the **Release build**, **CLI tests**, **MCP tests**, or **VS Code extension packaging**?

## Minimal credible matrix

```powershell
dotnet build-server shutdown
dotnet build Sbroenne.ExcelMcp.sln -c Release -p:NuGetAudit=false -nodeReuse:false
dotnet test tests\ExcelMcp.CLI.Tests\ExcelMcp.CLI.Tests.csproj -c Release --no-build
dotnet test tests\ExcelMcp.McpServer.Tests\ExcelMcp.McpServer.Tests.csproj -c Release --no-build
dotnet build-server shutdown
Set-Location vscode-extension
npm run package
```

## Why this exact shape

- `dotnet build-server shutdown` + `-nodeReuse:false` avoids false failures from locked build-task assemblies.
- `--no-build` keeps CLI and MCP test verdicts tied to the already-proven Release build.
- `npm run package` is the real extension gate; lighter checks can miss publish-time failures.
- Run the CLI and MCP validations **sequentially**, not in parallel. Concurrent .NET builds in this repo can poison the signal with locked assemblies or missing generated-reference artifacts.

## What to report

- Whether NuGet packages are actually outdated (`dotnet list ... package --outdated`)
- Whether npm devDependencies are outdated (`npm outdated --long`)
- Build failures separately from test failures
- CLI failures separately from MCP failures
- Packaging failures separately from compile/test failures
- Any evidence that the first failure poisoned later tests (for example: transport already configured, session not found, Excel process crashed)
