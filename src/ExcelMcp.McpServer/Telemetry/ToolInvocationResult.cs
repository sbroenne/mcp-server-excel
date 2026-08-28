namespace Sbroenne.ExcelMcp.McpServer.Telemetry;

internal enum ToolInvocationOutcome
{
    Succeeded,
    ExpectedNegative,
    Failed
}

internal enum ToolFailureClass
{
    InputState,
    ExternalDependency,
    TimeoutCancellation,
    ExcelRuntime,
    InternalProductFault,
    Unclassified
}

internal readonly record struct ToolInvocationResult(
    ToolInvocationOutcome Outcome,
    ToolFailureClass? FailureClass);
