---
applyTo: "src/ExcelMcp.Core/Commands/**/*.cs,src/ExcelMcp.Core/Models/Actions/**/*.cs,src/ExcelMcp.Service/**/*.cs,src/ExcelMcp.CLI/**/*.cs,src/ExcelMcp.McpServer/**/*.cs,src/ExcelMcp.Generators*/**/*.cs"
excludeAgent: "code-review"
---

# Generated contract checks

After changing a Core contract or generator, build Release and inspect generated
CLI options, batch JSON dispatch, Service routing, and MCP schemas. Names,
aliases, defaults, validation, results, and timeouts must agree. Hand-written MCP
tools may still own atomic no-session behavior, cancellation, or extra metadata.

Run `audit-core-coverage.ps1 -CheckNaming -FailOnGaps`,
`check-mcp-core-implementations.ps1`, and `check-doc-counts.ps1 -SkipBuild` under
`scripts` after that build. Compilation alone does not detect missing routes.
