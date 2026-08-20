---
applyTo: "src/**/*.cs"
excludeAgent: "code-review"
---

# Architecture patterns

## Layer boundaries

```text
MCP Server -> in-process ExcelMcpService -> Core commands -> Excel COM
CLI -> named-pipe daemon -> ExcelMcpService -> Core commands -> Excel COM
```

- `ExcelMcp.ComInterop` owns STA execution, COM lifetime, sessions, and shutdown.
- `ExcelMcp.Core` owns Excel behavior and typed result models.
- `ExcelMcp.Service` owns transport-neutral command dispatch.
- CLI and MCP are adapters. They must not reimplement Excel behavior.
- Source generators derive Service, CLI, and MCP routing from annotated Core
  interfaces. Change the source contract, not generated output.

## Core command pattern

Core methods are synchronous and accept `IExcelBatch`. Validate ordinary .NET
arguments before entering the batch, then perform Excel work inside
`batch.Execute`.

```csharp
public OperationResult Rename(IExcelBatch batch, string oldName, string newName)
{
    ArgumentException.ThrowIfNullOrWhiteSpace(oldName);
    ArgumentException.ThrowIfNullOrWhiteSpace(newName);

    return batch.Execute((ctx, ct) =>
    {
        Excel.Worksheet? sheet = null;
        try
        {
            ct.ThrowIfCancellationRequested();
            sheet = ComUtilities.FindSheet(ctx.Book, oldName)
                ?? throw new InvalidOperationException(
                    $"Worksheet '{oldName}' was not found.");
            sheet.Name = newName;
            return new OperationResult { Success = true };
        }
        finally
        {
            ComUtilities.Release(ref sheet);
        }
    });
}
```

- Do not wrap `batch.Execute` in a broad catch that returns another result.
- Acquire and release COM references inside the batch callback.
- Use cancellation checks in loops and before expensive operations.
- Reuse `ComUtilities` and feature helpers before adding new interop logic.

## Surface parity

For an action, parameter, default, or response change, trace this chain:

1. Core `[ServiceCategory]` interface and implementation
2. Generated Service command and argument model
3. CLI generated command/options and daemon routing
4. MCP action enum, mapping, schema, and generated/manual route
5. CLI, MCP, and Core tests
6. `skills/shared`, feature docs, and help text

Run `scripts\audit-core-coverage.ps1 -CheckNaming -FailOnGaps` and
`scripts\check-mcp-core-implementations.ps1` after contract changes.

## Code organization

- One public type per file; file name matches the type.
- Use partial classes to split large command implementations by domain.
- Keep validation close to the public contract and Excel behavior in Core.
- Prefer typed models over anonymous or loosely typed cross-layer payloads.
- Preserve established error categories and timeout/cancellation propagation.

## Performance and security

- Minimize COM round trips; use bulk range operations rather than per-cell loops.
- Reuse an open batch/session for related operations.
- Never return credentials or unsanitized connection strings.
- Keep destructive behavior explicit and consistent across CLI and MCP metadata.
