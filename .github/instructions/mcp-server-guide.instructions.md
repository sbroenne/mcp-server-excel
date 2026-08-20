---
applyTo: "src/ExcelMcp.McpServer/**/*.cs"
excludeAgent: "code-review"
---

# MCP Server development

MCP tool methods are static and synchronous. They route to the in-process
`ExcelMcpService`; do not add a named-pipe hop or duplicate Core behavior.

## Tool pattern

Use the existing generated registry and shared boundaries:

```csharp
public static string ExcelFeature(
    FeatureAction action,
    string? session_id = null,
    CancellationToken cancellationToken = default)
{
    using var cancellationScope =
        ExcelToolsBase.PushCancellationToken(cancellationToken);

    return ExcelToolsBase.ExecuteToolAction(
        "feature",
        ServiceRegistry.Feature.ToActionString(action),
        () => ServiceRegistry.Feature.RouteAction(
            action,
            session_id,
            ExcelToolsBase.ForwardToServiceFunc));
}
```

- `ServiceRegistry` and generated routes are preferred over hand-written action
  dispatch.
- Keep hand-written branches only for genuine special cases such as atomic
  no-session operations or extra MCP metadata.
- Pass the SDK cancellation token through `PushCancellationToken`.
- Use `ExcelToolsBase.ExecuteToolAction` for consistent telemetry and error
  serialization.

## Errors and JSON

- Tool execution failures return structured JSON with `success: false` and
  `isError: true`.
- Invalid arguments and unknown actions may throw the established argument or
  MCP protocol exception.
- Do not inspect a Core failure and throw a second generic exception; preserve
  the Service response's category, context, HRESULT, and retry information.
- Use shared `JsonOptions`; never construct JSON strings manually.
- Keep stdout exclusively for MCP JSON-RPC. Logs and diagnostics go to stderr.

## Schema and descriptions

- C# interface parameter names are camelCase; generated MCP names are snake_case.
- Keep action enums and `ToActionString` mappings exhaustive.
- Tool and parameter descriptions should state server-specific behavior,
  constraints, and disambiguation. Do not duplicate types or enum values already
  visible in the schema.
- Avoid emojis in XML documentation and generated LLM guidance.
- Keep destructive/read-only metadata accurate.

## Cross-surface verification

After an MCP contract change, verify the Core interface, generated Service route,
CLI command, MCP schema, tests, and `skills/shared` guidance. Run:

```powershell
dotnet build Sbroenne.ExcelMcp.sln -c Release
& .\scripts\audit-core-coverage.ps1 -CheckNaming -FailOnGaps
& .\scripts\check-mcp-core-implementations.ps1
```
