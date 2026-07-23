---
name: "mcp-schema-enum-regression"
description: "Keep generated MCP schemas strict-client compatible by avoiding nullable enum signatures."
domain: "mcp-server"
confidence: "high"
source: "observed"
---

## Context
Use this when an MCP client rejects tool schemas because enum arrays include invalid sentinel members (often surfaced as `null` or empty strings).

## Patterns
- Test the published contract with `ListToolsAsync()` and recurse every schema node that contains an `enum` array.
- Treat nullable enum parameters in MCP tool signatures as the hazard point; strict clients do not tolerate the SDK's sentinel output for them.
- Keep the primary `action` parameter as a required enum when you want discoverable action lists.
- Emit other action-specific enum inputs as optional strings, then parse them in generated MCP code before routing.
- For `[FromString]` enum paths, keep the string transport shape and let the existing service-side parser do the enum conversion.

## Examples
- `src/ExcelMcp.Generators.Mcp/McpToolGenerator.cs`: required `action` enum, optional enum-like inputs as strings, local parse for direct enum parameters.
- `tests/ExcelMcp.McpServer.Tests/Integration/Tools/GeneratedToolSchemaEnumRegressionTests.cs`: recursive schema regression plus focused optional-parameter check.

## Anti-Patterns
- Do not patch one generated `.g.cs` output by hand.
- Do not leave nullable enums in published MCP signatures and hope clients normalize them.
- Do not fix only `action`; scan every enum-bearing property across the tool catalog.
