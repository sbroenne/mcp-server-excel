# ADR-004: Generate entry-point contracts from Core definitions

**Status:** Current

## Context and decision

Annotated Core interfaces are the source for generated Service dispatch, CLI
commands, MCP tools, and operation metadata. Shared extraction connects those
outputs to the same definitions. Handwritten adapter code remains where a tool
owns behavior such as no-session work, cancellation, or image output.

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
- [Contract checks](agents/rules/coverage-prevention-strategy.md)
- [Agent-facing metadata sources](agents/rules/mcp-llm-guidance.md)
