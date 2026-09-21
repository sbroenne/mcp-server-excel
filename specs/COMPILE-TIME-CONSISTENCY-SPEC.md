# Compile-Time Consistency Contract

## Source of truth

Annotated Core interfaces define the shared operation contract:
`[ServiceCategory]` identifies a category, `[ServiceAction]` identifies an action,
and method signatures and parameter attributes define inputs and conversions.
`[McpTool]` supplies MCP naming, descriptions, and metadata.

Change those contracts or their generators, not emitted `.g.cs` files. The MCP
Server and CLI must expose matching behavior, defaults, validation, and results.
See the [system map](../CONTEXT.md) and [architecture](../docs/ARCHITECTURE.md).

## Generation map

| Source | Responsibility and output |
|--------|---------------------------|
| [Core attributes](../src/ExcelMcp.Core/Attributes/) and command interfaces | Category/action names, parameter contracts, and MCP tool metadata |
| [ExcelMcp.Generators.Shared](../src/ExcelMcp.Generators.Shared/) | Shared contract extraction and models used by generators |
| [ServiceRegistryGenerator](../src/ExcelMcp.Generators/ServiceRegistryGenerator.cs) | Generates Core's `ServiceRegistry.{Category}.g.cs` and `.Dispatch.g.cs`: action enums/constants, argument types, CLI settings/routing, MCP forwarding, and `DispatchToCore`; also emits shared dispatch helpers, contract and documentation manifests |
| [CliSettingsGenerator](../src/ExcelMcp.Generators.Cli/CliSettingsGenerator.cs) | Discovers annotated interfaces in referenced assemblies; generates CLI command classes and command registration |
| [McpToolGenerator](../src/ExcelMcp.Generators.Mcp/McpToolGenerator.cs) | Discovers annotated interfaces in referenced assemblies; generates MCP tool classes, parameter signatures/descriptions, `[McpServerTool]`, `[McpMeta]`, and routing |

CLI command names are derived from MCP tool names by removing underscores;
for example, `calculation_mode` becomes `calculationmode`. Service category
names are separate: both route calculation actions to `calculation.<action>`.

## Runtime and handwritten boundaries

[ExcelMcpService](../src/ExcelMcp.Service/ExcelMcpService.cs) owns session
management, selects category handlers and command instances, and wraps results.
Its category switch and session-specific orchestration are handwritten.
Per-action argument deserialization and calls into annotated Core commands use
generated `ServiceRegistry.{Category}.DispatchToCore()` methods.

The CLI command infrastructure and MCP transport helpers remain handwritten;
generation does not replace session lifecycle, error handling, or Excel COM
behavior. The MCP generator skips a tool name when a manual `[McpServerTool]`
implementation already exists, so those explicit exceptions still need review
when a contract changes.

Generation keeps action names, exposed parameters, and dispatch code derived
from the same interfaces. Compilation checks the generated calls against Core
signatures. It does **not** prove that handwritten category wiring, serialized
inputs, client schemas, or Excel behavior are correct.

## Change and validation procedure

1. Update the annotated Core contract and implementation together.
2. Follow the contract through Service dispatch, CLI, MCP, and shared guidance.
   Review handwritten handlers and manual-tool exceptions where applicable.
3. Add focused regression coverage for changed behavior. Existing
   [ServiceRegistry JSON parsing tests](../tests/ExcelMcp.Core.Tests/Unit/ServiceRegistryJsonParsingTests.cs)
   exercise Excel-independent conversion and dispatch; Excel operations require
   real Excel under the [testing policy](../docs/ADR-001-NO-UNIT-TESTS.md).
4. For runtime or generator changes, build the solution with zero warnings and
   run targeted tests and local Excel E2E as required by the
   [repository validation rules](../.github/copilot-instructions.md#build-and-validation).
   Run applicable existing coverage checks:
   - `scripts\audit-core-coverage.ps1 -CheckNaming -FailOnGaps`
   - `scripts\check-mcp-core-implementations.ps1`
   - `scripts\check-doc-counts.ps1` for advertised surface counts; use
     `-SkipBuild` only after a successful Release solution build in the worktree.

Documentation-only edits do not require synthetic tests or a build.
