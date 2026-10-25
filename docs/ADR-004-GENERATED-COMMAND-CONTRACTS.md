# ADR-004: Generate entry-point contracts from Core definitions

**Status:** Current

## Context and decision

Annotated Core interfaces are the source for generated Service dispatch, CLI
commands, MCP tools, and operation metadata. Shared extraction connects those
outputs to the same definitions.

The MCP Server keeps a hand-written tool layer for file and session lifecycle.
Worksheet tools, including transfers that do not require a session, and the
screenshot tool are generated from their Core definitions. A small MCP helper
packages screenshot results as image content; it does not declare or route tools.
All tools call the shared Service.

## Reasons and tradeoffs

Maintaining separate command inventories would make adding an operation easy
to miss in one entry point and allow names, defaults, or validation to diverge.
Generation makes common changes propagate from one definition.

The tradeoff is that a contract change can affect several outputs at once,
and compiling a generator does not prove the emitted public behavior is
correct. Intentional adapter exceptions still need their own coverage.
Generation reduces repeated implementation; it does not eliminate adapter
responsibilities.

## Implementation and guidance

- [Category annotation](../src/ExcelMcp.Core/Attributes/ServiceCategoryAttribute.cs)
- [Service generator](../src/ExcelMcp.Generators/ServiceRegistryGenerator.cs)
- [CLI generator](../src/ExcelMcp.Generators.Cli/CliSettingsGenerator.cs)
- [MCP generator](../src/ExcelMcp.Generators.Mcp/McpToolGenerator.cs)
- Hand-written MCP tool: [file](../src/ExcelMcp.McpServer/Tools/ExcelFileTool.cs)
- [Screenshot image conversion](../src/ExcelMcp.McpServer/Tools/ScreenshotImageResult.cs)
- [Contract checks](../src/ExcelMcp.Core/AGENTS.md)
- [Agent-facing metadata sources](../skills/AGENTS.md)
