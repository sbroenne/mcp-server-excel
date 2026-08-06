using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.ComInterop.ServiceClient;

/// <summary>
/// Protocol messages for CLI/MCP-to-service communication over named pipes.
/// Pattern: Client sends JSON request → Service executes → Returns JSON response.
/// All messages are newline-delimited JSON.
/// </summary>
public static class ServiceProtocol
{
    /// <summary>
    /// JSON serializer options for service protocol messages.
    /// </summary>
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>
    /// Serializes a message to JSON.
    /// </summary>
    public static string Serialize<T>(T message) => JsonSerializer.Serialize(message, JsonOptions);

    /// <summary>
    /// Deserializes a message from JSON.
    /// </summary>
    public static T? Deserialize<T>(string json) => JsonSerializer.Deserialize<T>(json, JsonOptions);
}

/// <summary>
/// Request sent from client (CLI or MCP) to service.
/// </summary>
public sealed class ServiceRequest
{
    /// <summary>Command to execute (e.g., "session.open", "sheet.list", "range.get-values").</summary>
    public required string Command { get; init; }

    /// <summary>Session ID for commands that operate on a session.</summary>
    public string? SessionId { get; init; }

    /// <summary>JSON-serialized command arguments.</summary>
    public string? Args { get; init; }

    /// <summary>Source of the request (CLI or MCP).</summary>
    public string? Source { get; init; }

    /// <summary>Return a mutation review plan without executing the command.</summary>
    public bool ReviewOnly { get; init; }

    /// <summary>Bound review identifier authorizing a previously reviewed mutation.</summary>
    public string? ReviewId { get; init; }

    /// <summary>Create a recoverable checkpoint immediately before a mutation.</summary>
    public bool Checkpoint { get; init; }

    /// <summary>
    /// Client-generated key that makes an executable session request safe to retry.
    /// An exact retry returns the original receipt; reusing the key for a different
    /// request fails closed.
    /// </summary>
    public string? IdempotencyKey { get; init; }
}

/// <summary>
/// Arguments for one ordered server-side batch request.
/// </summary>
public sealed class ServiceBatchRequest
{
    /// <summary>Operations to execute in their original order.</summary>
    public List<ServiceBatchOperation> Operations { get; init; } = [];

    /// <summary>
    /// Whether execution stops after the first failed operation. The wire-level
    /// default is fail-closed; clients that preserve continue-on-error behavior
    /// must send false explicitly.
    /// </summary>
    public bool StopOnError { get; init; } = true;
}

/// <summary>
/// Workflow-only plan envelope. This is intentionally separate from
/// <see cref="ServiceBatchRequest"/> so the legacy <c>session.batch</c> wire
/// contract remains unchanged.
/// </summary>
public sealed class WorkflowPlanRequest
{
    /// <summary>Ordered plan operations.</summary>
    public List<ServiceBatchOperation> Operations { get; init; } = [];
    /// <summary>Stop after the first failed operation.</summary>
    public bool StopOnError { get; init; } = true;
    /// <summary>Plan-level checkpoint policy.</summary>
    public WorkflowCheckpointMode CheckpointMode { get; init; } = WorkflowCheckpointMode.Inherit;
    /// <summary>
    /// Use one queued STA work item when the complete plan is compatible. Incompatible
    /// plans fall back to the normal sequential executor before any step is dispatched.
    /// </summary>
    public bool FastMode { get; init; } = true;
    /// <summary>
    /// Exact worksheet to inspect after the plan. Must be supplied together with
    /// <see cref="VerifyRangeAddress"/>.
    /// </summary>
    public string? VerifySheetName { get; init; }
    /// <summary>
    /// Exact rectangular range to inspect after the plan. Must be supplied together with
    /// <see cref="VerifySheetName"/>. The service never infers this from selection or UsedRange.
    /// </summary>
    public string? VerifyRangeAddress { get; init; }
}

/// <summary>Plan-level checkpoint policy for optimized workflow execution.</summary>
public enum WorkflowCheckpointMode
{
    /// <summary>Use the configured session policy.</summary>
    [JsonStringEnumMemberName("inherit")]
    Inherit,

    /// <summary>Do not create a plan checkpoint.</summary>
    [JsonStringEnumMemberName("off")]
    Off,

    /// <summary>Create one checkpoint before the first mutation.</summary>
    [JsonStringEnumMemberName("once")]
    Once,
}

/// <summary>Compact aggregate receipt outcome for one workflow plan.</summary>
public enum WorkflowPlanOutcome
{
    /// <summary>Every attempted step completed.</summary>
    [JsonStringEnumMemberName("completed")]
    Completed,

    /// <summary>The plan failed with a known outcome.</summary>
    [JsonStringEnumMemberName("failed")]
    Failed,

    /// <summary>The plan may have partially committed and must be reconciled.</summary>
    [JsonStringEnumMemberName("unknown")]
    Unknown,
}

/// <summary>Compact, replayable receipt for one workflow plan.</summary>
public sealed record WorkflowPlanReceipt
{
    /// <summary>Stable plan operation identifier.</summary>
    public string? PlanId { get; init; }
    /// <summary>Aggregate plan outcome.</summary>
    public WorkflowPlanOutcome Outcome { get; init; }
    /// <summary>Total requested operations.</summary>
    public int OperationCount { get; init; }
    /// <summary>Operations that started execution.</summary>
    public int AttemptedCount { get; init; }
    /// <summary>Operations that completed successfully.</summary>
    public int CompletedCount { get; init; }
    /// <summary>First failed or unknown operation index.</summary>
    public int? FailedIndex { get; init; }
    /// <summary>Shared plan checkpoint reference, when created.</summary>
    public WorkflowCheckpointReceipt? Checkpoint { get; init; }
    /// <summary>Compact ordered step statuses.</summary>
    public List<WorkflowStepReceipt> Steps { get; init; } = [];
    /// <summary>Execution path: fast, standard, or sequential-fallback.</summary>
    public string ExecutionMode { get; init; } = "standard";
    /// <summary>Whether the caller requested automatic fast execution.</summary>
    public bool FastModeRequested { get; init; }
    /// <summary>Whether the compatible one-STA executor was used.</summary>
    public bool FastModeUsed { get; init; }
    /// <summary>Why a requested fast plan used the sequential fallback.</summary>
    public string? FastModeFallbackReason { get; init; }
    /// <summary>Actual queued STA work items started while the plan ran, when available.</summary>
    public long? StaDispatchCount { get; init; }
    /// <summary>
    /// Bounded read-back of the caller-selected final range. Null when no verification
    /// scope was requested.
    /// </summary>
    public WorkflowRangeVerificationReceipt? Verification { get; init; }
}

/// <summary>
/// Compact, bounded verification of one explicit worksheet range after a workflow plan.
/// Counts and the fingerprint describe the inspected range; when <see cref="Status"/> is
/// <c>partiallyVerified</c>, <see cref="CellCount"/> remains the full requested size while
/// <see cref="InspectedCellCount"/> and <see cref="InspectedRangeAddress"/> disclose the sample.
/// </summary>
public sealed record WorkflowRangeVerificationReceipt
{
    /// <summary>verified, partiallyVerified, or notVerified.</summary>
    public string Status { get; init; } = "notVerified";
    /// <summary>Requested worksheet name.</summary>
    public string SheetName { get; init; } = string.Empty;
    /// <summary>Excel-canonical address of the full requested range when resolved.</summary>
    public string RangeAddress { get; init; } = string.Empty;
    /// <summary>Rows in the full requested range.</summary>
    public int? RowCount { get; init; }
    /// <summary>Columns in the full requested range.</summary>
    public int? ColumnCount { get; init; }
    /// <summary>Cells in the full requested range.</summary>
    public long? CellCount { get; init; }
    /// <summary>Cells actually inspected and included in the fingerprint.</summary>
    public int? InspectedCellCount { get; init; }
    /// <summary>Excel-canonical address of the inspected sample.</summary>
    public string? InspectedRangeAddress { get; init; }
    /// <summary>Non-empty cells in the inspected scope.</summary>
    public int? NonEmptyCellCount { get; init; }
    /// <summary>Formula cells in the inspected scope.</summary>
    public int? FormulaCellCount { get; init; }
    /// <summary>SHA-256 of the normalized values and formulas in the inspected scope.</summary>
    public string? Fingerprint { get; init; }
    /// <summary>At most two rows by four columns of normalized cell values.</summary>
    public IReadOnlyList<IReadOnlyList<object?>>? Preview { get; init; }
    /// <summary>Honest scope or failure limitation when verification was not complete.</summary>
    public string? Limitation { get; init; }
}

/// <summary>Compact checkpoint reference.</summary>
public sealed record WorkflowCheckpointReceipt
{
    /// <summary>Recovery identifier.</summary>
    public string RecoveryId { get; init; } = string.Empty;
    /// <summary>Checkpoint path relative to the safety root.</summary>
    public string RelativePath { get; init; } = string.Empty;
    /// <summary>SHA-256 of the checkpoint file.</summary>
    public string Sha256 { get; init; } = string.Empty;
    /// <summary>Checkpoint file size.</summary>
    public long Size { get; init; }
}

/// <summary>Compact status for one plan step.</summary>
public sealed record WorkflowStepReceipt
{
    /// <summary>Zero-based step index.</summary>
    public int Index { get; init; }
    /// <summary>Service command name.</summary>
    public string Command { get; init; } = string.Empty;
    /// <summary>Step status.</summary>
    public string Status { get; init; } = string.Empty;
    /// <summary>Structured error category, when failed.</summary>
    public string? ErrorCategory { get; init; }
}

/// <summary>
/// One operation inside a server-side batch.
/// </summary>
public sealed class ServiceBatchOperation
{
    /// <summary>Command to execute.</summary>
    public required string Command { get; init; }

    /// <summary>Optional session ID, which must match the enclosing request.</summary>
    public string? SessionId { get; init; }

    /// <summary>Command arguments as an embedded JSON value.</summary>
    public JsonElement? Args { get; init; }

    /// <summary>Return a review plan without executing this operation.</summary>
    public bool ReviewOnly { get; init; }

    /// <summary>Review identifier authorizing this operation.</summary>
    public string? ReviewId { get; init; }

    /// <summary>Create a checkpoint before this operation.</summary>
    public bool Checkpoint { get; init; }

    /// <summary>Client-generated retry key for this operation.</summary>
    public string? IdempotencyKey { get; init; }
}

/// <summary>
/// Ordered result envelope returned by a server-side batch.
/// </summary>
public sealed class ServiceBatchResponse
{
    /// <summary>Whether every attempted operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Whether every requested operation was attempted.</summary>
    public bool Completed { get; init; }

    /// <summary>Zero-based index of the first failed operation.</summary>
    public int? FailedIndex { get; init; }

    /// <summary>Results in the exact order in which operations were attempted.</summary>
    public List<ServiceBatchOperationResult> Results { get; init; } = [];
}

/// <summary>
/// Result for one attempted operation in a server-side batch.
/// </summary>
public sealed class ServiceBatchOperationResult
{
    /// <summary>Zero-based position in the request.</summary>
    public int Index { get; init; }

    /// <summary>
    /// Command that was attempted when needed for a standalone validation error.
    /// Normal ordered results omit this redundant copy; callers resolve it from Index.
    /// </summary>
    public string? Command { get; init; }

    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Command result as an embedded JSON value.</summary>
    public JsonElement? Result { get; init; }

    /// <summary>Error message when the operation failed.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Structured error category.</summary>
    public string? ErrorCategory { get; init; }

    /// <summary>Exception type when available.</summary>
    public string? ExceptionType { get; init; }

    /// <summary>HRESULT from a COM failure, when available.</summary>
    [JsonPropertyName("hresult")]
    public string? HResult { get; init; }
}

/// <summary>
/// Response sent from service to client.
/// </summary>
public sealed class ServiceResponse
{
    /// <summary>Whether the command succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>The service command that produced this response, when available.</summary>
    public string? Command { get; init; }

    /// <summary>The session ID associated with this response, when available.</summary>
    public string? SessionId { get; init; }

    /// <summary>Error message if Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Structured error category if Success is false.</summary>
    public string? ErrorCategory { get; init; }

    /// <summary>Exception type that produced the failure, when available.</summary>
    public string? ExceptionType { get; init; }

    /// <summary>HRESULT from a COM failure, when available.</summary>
    [JsonPropertyName("hresult")]
    public string? HResult { get; init; }

    /// <summary>Inner exception details, when available.</summary>
    public string? InnerError { get; init; }

    /// <summary>JSON-serialized result data.</summary>
    public string? Result { get; init; }
}


