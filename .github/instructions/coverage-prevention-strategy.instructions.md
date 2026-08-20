---
applyTo: "src/ExcelMcp.Core/Commands/**/*.cs,src/ExcelMcp.McpServer/**/*.cs"
excludeAgent: "code-review"
---

# Generated surface coverage

Core command interfaces annotated with `[ServiceCategory]` are the source
contracts for generated Service, CLI, and MCP routing. Do not follow old
hand-written "add a switch in every layer" patterns without first checking the
generator output.

## Contract change workflow

1. Add or change the method on the correct Core interface.
2. Implement the method in the matching Core command class.
3. Build Release so source generators refresh all generated surfaces.
4. Inspect generated diagnostics and the hand-written MCP tool only where that
   tool owns special routing, metadata, cancellation, or atomic behavior.
5. Verify CLI and MCP names, parameter defaults, validation, result shape, and
   timeout behavior remain identical.
6. Update focused tests and shared guidance.

## Required checks

```powershell
dotnet build Sbroenne.ExcelMcp.sln -c Release
& .\scripts\audit-core-coverage.ps1 -CheckNaming -FailOnGaps
& .\scripts\check-mcp-core-implementations.ps1
& .\scripts\check-doc-counts.ps1
```

The audits are authoritative. Do not compare hand-maintained method or action
counts in prose.

## Common incomplete changes

- Core method added without a generated service route
- MCP action added without the matching Core implementation
- Parameter renamed on one entry point only
- Default or validation changed without regenerating CLI/MCP schemas
- Hand-written tool exception not updated after a generated contract change
- Generated file edited instead of its interface, generator, or template
- Docs or skills updated with literal counts that disagree with the audits

Treat a build success as necessary but not sufficient; the coverage and naming
scripts catch contract gaps that compilation alone cannot.
