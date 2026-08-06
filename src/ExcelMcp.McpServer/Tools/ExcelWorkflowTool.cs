using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.ServiceClient;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// High-level Excel workflows that collapse a multi-step agent task into a small number of MCP calls.
/// </summary>
[McpServerToolType]
public static class ExcelWorkflowTool
{
    /// <summary>
    /// Executes an ordered Excel plan in one MCP call, or reports which optimized workflow features
    /// are available. Each plan operation uses the same command and arguments as the corresponding
    /// domain tool. Operations run in order and stop at the first error by default.
    /// </summary>
    /// <param name="action">Workflow action: capabilities, open-and-describe, or execute-plan.</param>
    /// <param name="session_id">Active session from file(open). Required for execute-plan.</param>
    /// <param name="file_path">Workbook path. Required for open-and-describe.</param>
    /// <param name="show">Whether to show Excel while opening the workbook.</param>
    /// <param name="timeout_seconds">Optional operation timeout in seconds.</param>
    /// <param name="preview_rows">Maximum preview rows per worksheet. Default 3; maximum 20.</param>
    /// <param name="preview_columns">Maximum preview columns per worksheet. Default 3; maximum 12.</param>
    /// <param name="operations">Ordered operations. Required for execute-plan.</param>
    /// <param name="stop_on_error">Stop at the first failed operation. Default true.</param>
    /// <param name="checkpoint_mode">Plan checkpoint policy: inherit, off, or once.</param>
    /// <param name="fast_mode">Automatically use one STA dispatch for compatible plans; incompatible plans fall back safely.</param>
    /// <param name="idempotency_key">Retry key for the whole plan.</param>
    /// <param name="verify_sheet_name">Exact worksheet to verify after the plan. Supply with verify_range_address.</param>
    /// <param name="verify_range_address">Exact rectangular range to verify after the plan. Supply with verify_sheet_name.</param>
    /// <param name="cancellationToken">Cancellation token supplied by the MCP host.</param>
    [McpServerTool(Name = "workflow", Title = "Optimized Excel Workflows", Destructive = true)]
    [McpMeta("category", "workflow")]
    [McpMeta("requiresSession", false)]
    [Description("Run optimized multi-step Excel workflows, inspect runtime capabilities, or open and summarize a workbook. Execute-plan supports one checkpoint, idempotent retries, and exact final-range verification.")]
    public static string ExcelWorkflow(
        WorkflowAction action,
        [DefaultValue(null)] string? session_id = null,
        [DefaultValue(null)] string? file_path = null,
        [DefaultValue(false)] bool show = false,
        [DefaultValue(null)] int? timeout_seconds = null,
        [DefaultValue(3)] int preview_rows = 3,
        [DefaultValue(3)] int preview_columns = 3,
        [DefaultValue(null)] List<WorkflowOperationInput>? operations = null,
        [DefaultValue(true)] bool stop_on_error = true,
        [DefaultValue(WorkflowCheckpointMode.Inherit)] WorkflowCheckpointMode checkpoint_mode = WorkflowCheckpointMode.Inherit,
        [DefaultValue(true)] bool fast_mode = true,
        [DefaultValue(null)] string? idempotency_key = null,
        [DefaultValue(null)] string? verify_sheet_name = null,
        [DefaultValue(null)] string? verify_range_address = null,
        CancellationToken cancellationToken = default)
    {
        using var cancellationScope = ExcelToolsBase.PushCancellationToken(cancellationToken);
        var actionName = action.ToActionString();

        return ExcelToolsBase.ExecuteToolAction(
            "workflow",
            actionName,
            () => action switch
            {
                WorkflowAction.Capabilities =>
                    GetCapabilities(),
                WorkflowAction.OpenAndDescribe => OpenAndDescribe(
                    file_path,
                    show,
                    timeout_seconds,
                    preview_rows,
                    preview_columns),
                WorkflowAction.ExecutePlan => ExecutePlan(
                    session_id,
                    operations,
                    stop_on_error,
                    checkpoint_mode,
                    fast_mode,
                    idempotency_key,
                    verify_sheet_name,
                    verify_range_address),
                _ => throw new ArgumentException($"Unknown workflow action: {action}", nameof(action)),
            });
    }

    private static string GetCapabilities()
    {
        return ExcelToolsBase.ForwardToServiceNoSession("workflow.capabilities");
    }

    private static string OpenAndDescribe(
        string? filePath,
        bool show,
        int? timeoutSeconds,
        int previewRows,
        int previewColumns)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("file_path is required for open-and-describe", nameof(filePath));
        }

        return ExcelToolsBase.ForwardToServiceNoSession(
            "workflow.open-and-describe",
            new
            {
                filePath,
                show,
                timeoutSeconds,
                previewRows,
                previewColumns,
            },
            timeoutSeconds);
    }

    private static string ExecutePlan(
        string? sessionId,
        List<WorkflowOperationInput>? operations,
        bool stopOnError,
        WorkflowCheckpointMode checkpointMode,
        bool fastMode,
        string? idempotencyKey,
        string? verifySheetName,
        string? verifyRangeAddress)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            throw new ArgumentException("session_id is required for execute-plan", nameof(sessionId));
        }

        if (operations is not { Count: > 0 })
        {
            throw new ArgumentException("operations must contain at least one operation for execute-plan", nameof(operations));
        }

        bool hasVerifySheet = !string.IsNullOrWhiteSpace(verifySheetName);
        bool hasVerifyRange = !string.IsNullOrWhiteSpace(verifyRangeAddress);
        if (hasVerifySheet != hasVerifyRange)
        {
            throw new ArgumentException(
                "verify_sheet_name and verify_range_address must be supplied together for execute-plan");
        }

        using var safetyScope = ExcelToolsBase.PushSafetyOptions(
            false,
            null,
            checkpointMode == WorkflowCheckpointMode.Once,
            idempotencyKey);

        return ExcelToolsBase.ForwardToService(
            "workflow.execute-plan",
            sessionId,
            new
            {
                operations,
                stopOnError,
                checkpointMode,
                fastMode,
                verifySheetName = hasVerifySheet ? verifySheetName!.Trim() : null,
                verifyRangeAddress = hasVerifyRange ? verifyRangeAddress!.Trim() : null,
            });
    }
}

/// <summary>
/// One command inside an optimized workflow plan.
/// </summary>
public sealed class WorkflowOperationInput
{
    /// <summary>Service command, for example range.set-values or rangeformat.format-range.</summary>
    [Description("Service command, for example range.set-values or rangeformat.format-range")]
    public required string Command { get; init; }

    /// <summary>Arguments for the command as a JSON object keyed by the service parameter names.</summary>
    [Description("JSON object keyed by the selected service command's parameter names")]
    public Dictionary<string, JsonElement>? Args { get; init; }

    /// <summary>
    /// Forbidden compatibility field. Workflow v2 plan review is unavailable;
    /// this value must remain false.
    /// </summary>
    public bool ReviewOnly { get; init; }

    /// <summary>
    /// Forbidden compatibility field. Workflow v2 plan review is unavailable;
    /// this value must remain null.
    /// </summary>
    public string? ReviewId { get; init; }

    /// <summary>
    /// Legacy session.batch field; workflow plans use checkpoint_mode once/off/inherit.
    /// True is rejected by workflow.execute-plan.
    /// </summary>
    public bool Checkpoint { get; init; }

    /// <summary>
    /// Legacy session.batch field; workflow plans use the top-level idempotency_key.
    /// Non-null values are rejected by workflow.execute-plan.
    /// </summary>
    public string? IdempotencyKey { get; init; }
}
