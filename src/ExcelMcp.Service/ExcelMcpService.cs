using System.Collections.Concurrent;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using ServiceBatchOperation = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceBatchOperation;
using ServiceBatchOperationResult = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceBatchOperationResult;
using ServiceBatchRequest = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceBatchRequest;
using ServiceBatchResponse = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceBatchResponse;
using Sbroenne.ExcelMcp.ComInterop.ServiceClient;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.Diag;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.PythonInExcel;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.ReportFormat;
using Sbroenne.ExcelMcp.Core.Commands.Outline;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Service.Idempotency;
using Sbroenne.ExcelMcp.Service.Safety;
using Sbroenne.ExcelMcp.Service.Workflow;
using StreamJsonRpc;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Commands.Window;
using Sbroenne.ExcelMcp.Generated;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// The ExcelMCP Service. Holds SessionManager and executes Core commands.
/// Runs in-process within the host (MCP Server or CLI), accepting commands via named pipe.
/// The named pipe enables cross-thread communication between the host's request threads
/// and the service's STA thread (required for COM interop).
/// </summary>
public sealed class ExcelMcpService : IDisposable
{
    private const int MaximumBatchOperations = 256;
    private readonly SessionManager _sessionManager = new();
    private readonly ConcurrentDictionary<string, byte> _knownSessionIds = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<Task, byte> _activeConnectionTasks = new();
    private readonly CancellationTokenSource _shutdownCts = new();
    private readonly DateTime _startTime = DateTime.UtcNow;
    private string _pipeName = "";
    private TimeSpan? _idleTimeout;
    private DateTime _lastActivityTime = DateTime.UtcNow;
    private bool _disposed;

    // Core command instances - use concrete types per CA1859
    private readonly RangeCommands _rangeCommands = new();
    private readonly SheetCommands _sheetCommands = new();
    private readonly TableCommands _tableCommands = new();
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly PivotTableCommands _pivotTableCommands = new();
    private readonly SlicerCommands _slicerCommands = new();
    private readonly ChartCommands _chartCommands = new();
    private readonly ConnectionCommands _connectionCommands = new();
    private readonly NamedRangeCommands _namedRangeCommands = new();
    private readonly ConditionalFormattingCommands _conditionalFormatCommands = new();
    private readonly VbaCommands _vbaCommands = new();
    private readonly DataModelCommands _dataModelCommands = new();
    private readonly CalculationModeCommands _calculationModeCommands = new();
    private readonly ScreenshotCommands _screenshotCommands = new();
    private readonly DiagCommands _diagCommands = new();
    private readonly WindowCommands _windowCommands = new();
    private readonly PythonInExcelCommands _pythonInExcelCommands = new();
    private readonly ReportFormatCommands _reportFormatCommands = new();
    private readonly OutlineCommands _outlineCommands = new();
    private readonly WorkbookSafetyCoordinator _safetyCoordinator;
    private readonly IdempotencyCoordinator _idempotencyCoordinator = new();
    private readonly WorkflowRuntimeManifest _workflowRuntimeManifest;

    public ExcelMcpService() : this(safetyStateRoot: null, workflowRuntimeManifest: null)
    {
    }

    /// <summary>
    /// Creates the service with an optional durable safety-state root.
    /// </summary>
    public ExcelMcpService(string? safetyStateRoot) : this(safetyStateRoot, workflowRuntimeManifest: null)
    {
    }

    /// <summary>
    /// Creates the service with a host-owned workflow runtime manifest. Hosts that expose
    /// <c>workflow.capabilities</c> should pass their own executable assembly to the manifest.
    /// </summary>
    public ExcelMcpService(WorkflowRuntimeManifest workflowRuntimeManifest)
        : this(safetyStateRoot: null, workflowRuntimeManifest: workflowRuntimeManifest)
    {
    }

    /// <summary>Creates the service with optional safety state and host runtime identity.</summary>
    public ExcelMcpService(string? safetyStateRoot, WorkflowRuntimeManifest? workflowRuntimeManifest)
    {
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
        _safetyCoordinator = new WorkbookSafetyCoordinator(safetyStateRoot);
        _workflowRuntimeManifest = workflowRuntimeManifest ?? WorkflowRuntimeManifest.CreateServiceDefault();
        _sessionManager.DeadSessionCleanupStarting += HandleDeadSessionCleanupStarting;
    }

    public DateTime StartTime => _startTime;
    public int SessionCount => _sessionManager.GetActiveSessions().Count;
    public SessionManager SessionManager => _sessionManager;

    /// <summary>
    /// Runs the service in-process, listening for commands on the named pipe.
    /// This method blocks until shutdown is requested via <see cref="RequestShutdown"/>.
    /// </summary>
    /// <param name="pipeName">The named pipe to listen on.</param>
    /// <param name="idleTimeout">Optional idle timeout. Service shuts down after this duration with no active sessions. Null = no timeout.</param>
    public async Task RunAsync(string pipeName, TimeSpan? idleTimeout = null)
    {
        _pipeName = pipeName;
        _idleTimeout = idleTimeout;
        await RunPipeServerAsync(_shutdownCts.Token);
    }

    public void RequestShutdown() => _shutdownCts.Cancel();

    private void RequestShutdownAfterResponse()
    {
        _ = Task.Run(async () =>
        {
            await Task.Delay(100);
            RequestShutdown();
        });
    }

    // Exposed for testing — backoff parameters for pipe server accept loop error recovery
    internal static readonly TimeSpan InitialBackoff = TimeSpan.FromMilliseconds(100);
    internal static readonly TimeSpan MaxBackoff = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Records client activity to keep the idle timeout monitor alive.
    /// Called by <see cref="Rpc.DaemonRpcTarget"/> on each incoming RPC call.
    /// </summary>
    internal void RecordActivity() => _lastActivityTime = DateTime.UtcNow;

    private async Task RunPipeServerAsync(CancellationToken cancellationToken)
    {
        // Use a semaphore to limit concurrent connections (prevents resource exhaustion)
        using var connectionLimit = new SemaphoreSlim(10, 10);

        // Start idle timeout monitor if configured
        if (_idleTimeout.HasValue)
        {
            _ = Task.Run(() => MonitorIdleTimeoutAsync(cancellationToken), cancellationToken);
        }

        var currentBackoff = InitialBackoff;

        while (!cancellationToken.IsCancellationRequested)
        {
            NamedPipeServerStream? server = null;
            try
            {
                server = ServiceSecurity.CreateSecureServer(_pipeName);
                await server.WaitForConnectionAsync(cancellationToken);

                // Success — reset backoff
                currentBackoff = InitialBackoff;

                // Record activity on each connection
                _lastActivityTime = DateTime.UtcNow;

                // Capture server for the task
                var clientServer = server;
                server = null; // Prevent disposal in finally - task owns it now

                var connectionTask = Task.Run(async () =>
                {
                    await connectionLimit.WaitAsync();
                    try
                    {
                        var rpcTarget = new DaemonRpcTarget(this);
                        using var rpc = JsonRpc.Attach(clientServer, rpcTarget);
                        await rpc.Completion; // Waits until client disconnects
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException)
                    {
                        System.Diagnostics.Debug.WriteLine($"RPC connection failed: {ex.Message}");
                    }
                    catch (OperationCanceledException ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"RPC connection cancelled: {ex.Message}");
                    }
                    finally
                    {
                        connectionLimit.Release();
                        try { if (clientServer.IsConnected) clientServer.Disconnect(); }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Pipe disconnect cleanup failed: {ex.Message}");
                        }

                        try { await clientServer.DisposeAsync(); }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Pipe disposal cleanup failed: {ex.Message}");
                        }
                    }
                });
                _activeConnectionTasks.TryAdd(connectionTask, 0);
                _ = connectionTask.ContinueWith(
                    completed => _activeConnectionTasks.TryRemove(completed, out _),
                    CancellationToken.None,
                    TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception)
            {
                // Backoff to prevent CPU spin when errors repeat (e.g. pipe creation failure).
                // Doubles each iteration: 100ms → 200ms → 400ms → … → 5s cap.
                // Resets to 100ms on next successful connection.
                try { await Task.Delay(currentBackoff, cancellationToken); } catch (OperationCanceledException) { break; }
                currentBackoff = TimeSpan.FromMilliseconds(Math.Min(currentBackoff.TotalMilliseconds * 2, MaxBackoff.TotalMilliseconds));
            }
            finally
            {
                if (server != null)
                {
                    try { if (server.IsConnected) server.Disconnect(); } catch (Exception) { /* Cleanup — disconnect may fail if client already disconnected */ }
                    await server.DisposeAsync();
                }
            }
        }

        if (!_activeConnectionTasks.IsEmpty)
        {
            await Task.WhenAll(_activeConnectionTasks.Keys.Select(ObserveConnectionTaskAsync));
        }
    }

    private static async Task ObserveConnectionTaskAsync(Task connectionTask)
    {
        try
        {
            await connectionTask;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            System.Diagnostics.Debug.WriteLine($"RPC connection drain failed: {ex.Message}");
        }
        catch (OperationCanceledException ex)
        {
            System.Diagnostics.Debug.WriteLine($"RPC connection drain cancelled: {ex.Message}");
        }
    }

    private async Task MonitorIdleTimeoutAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            await Task.Delay(TimeSpan.FromSeconds(30), cancellationToken);

            var hasSessions = _sessionManager.GetActiveSessions().Count > 0;
            if (hasSessions)
            {
                _lastActivityTime = DateTime.UtcNow;
                continue;
            }

            var idleTime = DateTime.UtcNow - _lastActivityTime;
            if (idleTime >= _idleTimeout!.Value)
            {
                RequestShutdown();
                break;
            }
        }
    }

    /// <summary>
    /// Processes a service request directly (in-process, no pipe).
    /// Used by the MCP Server for direct in-process communication.
    /// </summary>
    public Task<ServiceResponse> ProcessAsync(ServiceRequest request) =>
        _idempotencyCoordinator.ExecuteAsync(request, () => ProcessCoreAsync(request));

    private async Task<ServiceResponse> ProcessCoreAsync(ServiceRequest request)
    {
        try
        {
            // Route command
            var parts = request.Command.Split('.', 2);
            var category = parts[0];
            var action = parts.Length > 1 ? parts[1] : "";

            ServiceResponse response = category switch
            {
                "service" => HandleServiceCommand(action),
                "workflow" => await HandleWorkflowCommandAsync(action, request),
                "session" => string.Equals(action, "batch", StringComparison.Ordinal)
                    ? await HandleSessionBatchAsync(request)
                    : HandleSessionCommand(action, request),
                "recovery" => HandleRecoveryCommand(action, request),
                "sheet" or "sheetstyle" => await DispatchSheetAsync(action, request),
                "range" or "rangeedit" or "rangeformat" or "rangelink" => await DispatchRangeAsync(action, request),
                "table" or "tablecolumn" => await DispatchTableAsync(action, request),
                "powerquery" => await DispatchSimpleAsync<PowerQueryAction>(action, request,
                    ServiceRegistry.PowerQuery.TryParseAction,
                    (a, batch) => ServiceRegistry.PowerQuery.DispatchToCore(_powerQueryCommands, a, batch, request.Args)),
                "pivottable" => await DispatchSimpleAsync<PivotTableAction>(action, request,
                    ServiceRegistry.PivotTable.TryParseAction,
                    (a, batch) => ServiceRegistry.PivotTable.DispatchToCore(_pivotTableCommands, a, batch, request.Args)),
                "pivottablefield" => await DispatchSimpleAsync<PivotTableFieldAction>(action, request,
                    ServiceRegistry.PivotTableField.TryParseAction,
                    (a, batch) => ServiceRegistry.PivotTableField.DispatchToCore(_pivotTableCommands, a, batch, request.Args)),
                "pivottablecalc" => await DispatchSimpleAsync<PivotTableCalcAction>(action, request,
                    ServiceRegistry.PivotTableCalc.TryParseAction,
                    (a, batch) => ServiceRegistry.PivotTableCalc.DispatchToCore(_pivotTableCommands, a, batch, request.Args)),
                "chart" => await DispatchSimpleAsync<ChartAction>(action, request,
                    ServiceRegistry.Chart.TryParseAction,
                    (a, batch) => ServiceRegistry.Chart.DispatchToCore(_chartCommands, a, batch, request.Args)),
                "chartconfig" => await DispatchSimpleAsync<ChartConfigAction>(action, request,
                    ServiceRegistry.ChartConfig.TryParseAction,
                    (a, batch) => ServiceRegistry.ChartConfig.DispatchToCore(_chartCommands, a, batch, request.Args)),
                "connection" => await DispatchSimpleAsync<ConnectionAction>(action, request,
                    ServiceRegistry.Connection.TryParseAction,
                    (a, batch) => ServiceRegistry.Connection.DispatchToCore(_connectionCommands, a, batch, request.Args)),
                "calculation" => await DispatchSimpleAsync<CalculationAction>(action, request,
                    ServiceRegistry.Calculation.TryParseAction,
                    (a, batch) => ServiceRegistry.Calculation.DispatchToCore(_calculationModeCommands, a, batch, request.Args)),
                "namedrange" => await DispatchSimpleAsync<NamedRangeAction>(action, request,
                    ServiceRegistry.NamedRange.TryParseAction,
                    (a, batch) => ServiceRegistry.NamedRange.DispatchToCore(_namedRangeCommands, a, batch, request.Args)),
                "conditionalformat" => await DispatchSimpleAsync<ConditionalFormatAction>(action, request,
                    ServiceRegistry.ConditionalFormat.TryParseAction,
                    (a, batch) => ServiceRegistry.ConditionalFormat.DispatchToCore(_conditionalFormatCommands, a, batch, request.Args)),
                "vba" => await DispatchSimpleAsync<VbaAction>(action, request,
                    ServiceRegistry.Vba.TryParseAction,
                    (a, batch) => ServiceRegistry.Vba.DispatchToCore(_vbaCommands, a, batch, request.Args)),
                "datamodel" => await DispatchSimpleAsync<DataModelAction>(action, request,
                    ServiceRegistry.DataModel.TryParseAction,
                    (a, batch) => ServiceRegistry.DataModel.DispatchToCore(_dataModelCommands, a, batch, request.Args)),
                "datamodelrel" => await DispatchSimpleAsync<DataModelRelAction>(action, request,
                    ServiceRegistry.DataModelRel.TryParseAction,
                    (a, batch) => ServiceRegistry.DataModelRel.DispatchToCore(_dataModelCommands, a, batch, request.Args)),
                "slicer" => await DispatchSimpleAsync<SlicerAction>(action, request,
                    ServiceRegistry.Slicer.TryParseAction,
                    (a, batch) => ServiceRegistry.Slicer.DispatchToCore(_slicerCommands, a, batch, request.Args)),
                "screenshot" => await DispatchSimpleAsync<ScreenshotAction>(action, request,
                    ServiceRegistry.Screenshot.TryParseAction,
                    (a, batch) => ServiceRegistry.Screenshot.DispatchToCore(_screenshotCommands, a, batch, request.Args)),
                "window" => await DispatchWindowAsync(action, request),
                "reportformat" => await DispatchSimpleAsync<ReportFormatAction>(action, request,
                    ServiceRegistry.ReportFormat.TryParseAction,
                    (a, batch) => ServiceRegistry.ReportFormat.DispatchToCore(_reportFormatCommands, a, batch, request.Args)),
                "outline" => await DispatchSimpleAsync<OutlineAction>(action, request,
                    ServiceRegistry.Outline.TryParseAction,
                    (a, batch) => ServiceRegistry.Outline.DispatchToCore(_outlineCommands, a, batch, request.Args)),
                "diag" => DispatchSessionless(action, request),
                "pythoninexcel" => await DispatchSimpleAsync<PythonInExcelAction>(action, request,
                    ServiceRegistry.PythonInExcel.TryParseAction,
                    (a, batch) => ServiceRegistry.PythonInExcel.DispatchToCore(_pythonInExcelCommands, a, batch, request.Args)),
                _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown command category: {category}" }
            };

            return AttachRequestContext(request, response);
        }
        catch (Exception ex)
        {
            // Include type name so callers can distinguish exception kinds (GitHub #482, Bug 5)
            return CreateErrorResponse(ex, request.Command, request.SessionId);
        }
    }

    // === SERVICE COMMANDS ===

    private ServiceResponse HandleServiceCommand(string action)
    {
        return action switch
        {
            "ping" => new ServiceResponse { Success = true },
            "shutdown" => HandleShutdown(),
            "status" => HandleStatus(),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown service action: {action}" }
        };
    }

    private ServiceResponse HandleShutdown()
    {
        RequestShutdownAfterResponse();
        return new ServiceResponse { Success = true };
    }

    private ServiceResponse HandleStatus()
    {
        var status = new ServiceStatus
        {
            Running = true,
            ProcessId = Environment.ProcessId,
            SessionCount = _sessionManager.GetActiveSessions().Count,
            StartTime = _startTime
        };
        return new ServiceResponse { Success = true, Result = JsonSerializer.Serialize(status, ServiceProtocol.JsonOptions) };
    }

    // === SESSION COMMANDS ===

    private Task<ServiceResponse> HandleWorkflowCommandAsync(string action, ServiceRequest request)
    {
        return action switch
        {
            "capabilities" => Task.FromResult(new ServiceResponse
            {
                Success = true,
                Result = ServiceProtocol.Serialize(new
                {
                    success = true,
                    workflowInterfaceVersion = WorkflowRuntimeManifest.InterfaceVersion,
                    runtimeHost = _workflowRuntimeManifest.ServerName,
                    serverVersion = _workflowRuntimeManifest.ServerVersion,
                    buildFingerprint = _workflowRuntimeManifest.BuildFingerprint,
                    toolProfile = _workflowRuntimeManifest.ToolProfile,
                    toolProfileVersion = _workflowRuntimeManifest.ToolProfileVersion,
                    toolProfileFallback = _workflowRuntimeManifest.ToolProfileFallback,
                    toolProfileTools = _workflowRuntimeManifest.ToolProfileTools,
                    toolProfileManifestHash = _workflowRuntimeManifest.ToolProfileManifestHash,
                    executePlan = true,
                    openAndDescribe = true,
                    fastMode = true,
                    fastModeVersion = "1",
                    fastModeFallback = "sequential",
                    fastModeCompatibleCategories = WorkflowFastPathPolicy.CompatibleCategories,
                    fastModeReceipt = true,
                    compactReceipts = true,
                    planCheckpoint = true,
                    planIdempotency = true,
                    finalRangeVerification = true,
                    planReview = false,
                }),
            }),
            "open-and-describe" => Task.FromResult(HandleWorkflowOpenAndDescribe(request)),
            "execute-plan" => HandleWorkflowPlanAsync(request),
            _ => Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Unknown workflow action: {action}",
            }),
        };
    }

    private ServiceResponse HandleWorkflowOpenAndDescribe(ServiceRequest request)
    {
        WorkflowOpenAndDescribeArgs? args;
        try
        {
            args = ServiceProtocol.Deserialize<WorkflowOpenAndDescribeArgs>(request.Args ?? "{}");
        }
        catch (JsonException ex)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Invalid workflow.open-and-describe arguments: {ex.Message}",
            };
        }

        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "filePath is required",
            };
        }

        int previewRows = args.PreviewRows ?? WorkbookManifestReader.DefaultPreviewRows;
        int previewColumns = args.PreviewColumns ?? WorkbookManifestReader.DefaultPreviewColumns;
        if (previewRows is < 1 or > WorkbookManifestReader.MaximumPreviewRows)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"previewRows must be between 1 and {WorkbookManifestReader.MaximumPreviewRows}",
            };
        }

        if (previewColumns is < 1 or > WorkbookManifestReader.MaximumPreviewColumns)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"previewColumns must be between 1 and {WorkbookManifestReader.MaximumPreviewColumns}",
            };
        }

        if (args.TimeoutSeconds is <= 0)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "timeoutSeconds must be greater than zero",
            };
        }

        string? sessionId = null;
        try
        {
            TimeSpan? timeout = args.TimeoutSeconds.HasValue
                ? TimeSpan.FromSeconds(args.TimeoutSeconds.Value)
                : null;
            sessionId = _sessionManager.CreateSession(
                args.FilePath,
                show: args.Show,
                operationTimeout: timeout,
                origin: SessionOrigin.MCP);

            var sessionError = TryBeginUsableSession(sessionId, out var batch);
            if (sessionError != null)
            {
                CleanupOwnedWorkflowSession(sessionId);
                return sessionError;
            }

            WorkbookManifest manifest;
            try
            {
                manifest = WorkbookManifestReader.Read(batch!, sessionId, previewRows, previewColumns);
            }
            finally
            {
                _sessionManager.EndOperation(sessionId);
            }

            _knownSessionIds.TryAdd(sessionId, 0);
            return new ServiceResponse
            {
                Success = true,
                SessionId = sessionId,
                Result = ServiceProtocol.Serialize(manifest),
            };
        }
        catch (Exception ex)
        {
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                CleanupOwnedWorkflowSession(sessionId);
            }

            return CreateErrorResponse(ex, request.Command);
        }
    }

    private async Task<ServiceResponse> HandleWorkflowPlanAsync(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorCategory = "PlanNotExecuted", ErrorMessage = "sessionId is required" };
        }

        if (request.ReviewOnly)
        {
            return new ServiceResponse { Success = false, ErrorCategory = "PlanReviewUnavailable", ErrorMessage = "Plan-level review is not available for workflow.execute-plan yet." };
        }

        if (!string.IsNullOrWhiteSpace(request.ReviewId))
        {
            return new ServiceResponse { Success = false, ErrorCategory = "PlanReviewUnavailable", ErrorMessage = "Plan-level review is not available for workflow.execute-plan yet." };
        }

        WorkflowPlanRequest? plan;
        try
        {
            plan = ServiceProtocol.Deserialize<WorkflowPlanRequest>(request.Args ?? "{}");
        }
        catch (JsonException ex)
        {
            return new ServiceResponse { Success = false, ErrorCategory = "PlanNotExecuted", ErrorMessage = $"Invalid workflow.execute-plan arguments: {ex.Message}" };
        }

        if (plan?.Operations is not { Count: > 0 })
        {
            return new ServiceResponse { Success = false, ErrorCategory = "PlanNotExecuted", ErrorMessage = "workflow.execute-plan requires at least one operation" };
        }

        if (plan.Operations.Count > MaximumBatchOperations)
        {
            return new ServiceResponse { Success = false, ErrorCategory = "PlanNotExecuted", ErrorMessage = $"workflow.execute-plan accepts at most {MaximumBatchOperations} operations" };
        }

        bool hasVerificationSheet = !string.IsNullOrWhiteSpace(plan.VerifySheetName);
        bool hasVerificationRange = !string.IsNullOrWhiteSpace(plan.VerifyRangeAddress);
        if (hasVerificationSheet != hasVerificationRange)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "PlanNotExecuted",
                ErrorMessage = "verifySheetName and verifyRangeAddress must be supplied together",
            };
        }

        string? verificationSheetName = hasVerificationSheet ? plan.VerifySheetName!.Trim() : null;
        string? verificationRangeAddress = hasVerificationRange ? plan.VerifyRangeAddress!.Trim() : null;
        bool verificationRequested = verificationSheetName is not null;

        var mutationSteps = new bool[plan.Operations.Count];
        bool hasMutation = false;
        for (int index = 0; index < plan.Operations.Count; index++)
        {
            var operation = plan.Operations[index];
            string? validationError = ValidateBatchOperation(request.SessionId, operation);
            if (validationError is not null)
            {
                return new ServiceResponse { Success = false, ErrorCategory = "PlanNotExecuted", ErrorMessage = $"Operation {index}: {validationError}" };
            }

            if (operation.ReviewOnly || operation.ReviewId is not null || operation.Checkpoint || operation.IdempotencyKey is not null)
            {
                return new ServiceResponse { Success = false, ErrorCategory = "PlanSafetyConflict", ErrorMessage = $"Operation {index} contains per-step safety options; use plan-level options instead." };
            }

            var descriptor = ServiceRegistry.GetSafetyDescriptor(operation.Command);
            if (!descriptor.ExplicitlyClassified)
            {
                return new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "PlanSafetyConflict",
                    ErrorMessage = $"Operation {index} command '{operation.Command}' has no explicit safety classification and cannot run inside an optimized plan."
                };
            }

            mutationSteps[index] = descriptor.IsMutation;
            hasMutation |= descriptor.IsMutation;
        }

        var safetyConfiguration = _safetyCoordinator.GetConfiguration(request.SessionId);
        if (hasMutation && safetyConfiguration.ReviewMode == ReviewMode.Required)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "PlanReviewUnavailable",
                ErrorMessage = "This session requires mutation review, but plan-level review is not available for workflow.execute-plan yet."
            };
        }

        if (hasMutation && plan.CheckpointMode == WorkflowCheckpointMode.Off &&
            safetyConfiguration.CheckpointMode == CheckpointMode.Required)
        {
            return new ServiceResponse { Success = false, ErrorCategory = "PlanOptionConflict", ErrorMessage = "checkpoint_mode=off cannot bypass the required session checkpoint policy." };
        }

        bool sharedCheckpointRequested = hasMutation &&
            (plan.CheckpointMode == WorkflowCheckpointMode.Once ||
             (plan.CheckpointMode == WorkflowCheckpointMode.Inherit &&
              (request.Checkpoint || safetyConfiguration.CheckpointMode == CheckpointMode.Required)));

        string planId = Guid.NewGuid().ToString("N");
        DateTime planStartedAtUtc = DateTime.UtcNow;
        var steps = new List<WorkflowStepReceipt>(plan.Operations.Count);
        int attempted = 0;
        int completed = 0;
        int completedMutations = 0;
        int? failedIndex = null;
        ServiceResponse? firstFailure = null;
        bool checkpointAttempted = false;
        bool knownCheckpointFailure = false;
        bool unknownOutcomeSeen = false;
        IDisposable? requiredCheckpointScope = null;
        WorkflowCheckpointReceipt? checkpointReceipt = null;
        WorkflowRangeVerificationReceipt? verificationReceipt = null;
        string? fastModeFallbackReason = plan.FastMode
            ? WorkflowFastPathPolicy.GetFallbackReason(plan, safetyConfiguration, sharedCheckpointRequested)
            : null;
        bool fastModeUsed = plan.FastMode && fastModeFallbackReason is null;
        string executionMode = fastModeUsed
            ? "fast"
            : plan.FastMode ? "sequential-fallback" : "standard";
        long? staDispatchCount = null;

        ServiceRequest CreateOperationRequest(ServiceBatchOperation operation, bool requestSharedCheckpoint) => new()
        {
            Command = operation.Command,
            SessionId = request.SessionId,
            Args = GetOperationArgsJson(operation.Args),
            Source = request.Source,
            ReviewOnly = false,
            ReviewId = null,
            Checkpoint = requestSharedCheckpoint,
        };

        bool RecordOperationResponse(
            int index,
            ServiceBatchOperation operation,
            bool isMutation,
            bool requestSharedCheckpoint,
            ServiceResponse originalResponse)
        {
            var response = originalResponse;
            bool terminalCheckpointFailure = false;
            if (requestSharedCheckpoint)
            {
                checkpointReceipt = TryGetCheckpointReceipt(response.Result) ??
                    TryGetLatestCheckpointReceipt(request.SessionId, planStartedAtUtc, operation.Command);
                if (checkpointReceipt is not null)
                {
                    // A valid shared checkpoint now protects every later mutation in
                    // this lease, including plans that continue after a known failure.
                    requiredCheckpointScope ??= WorkbookSafetyCoordinator.SuppressRequiredCheckpoints(request.SessionId);
                }
                else if (string.Equals(response.ErrorCategory, "CheckpointFailed", StringComparison.Ordinal))
                {
                    knownCheckpointFailure = true;
                    terminalCheckpointFailure = true;
                }
                else
                {
                    response = new ServiceResponse
                    {
                        Success = false,
                        Command = operation.Command,
                        SessionId = request.SessionId,
                        ErrorCategory = "JournalPersistenceFailed",
                        ErrorMessage = "The shared checkpoint could not be verified after dispatch. The workbook may have changed; reconcile it before retrying.",
                    };
                    terminalCheckpointFailure = true;
                }
            }

            if (response.Success)
            {
                completed++;
                if (isMutation)
                {
                    completedMutations++;
                }
                steps.Add(new WorkflowStepReceipt { Index = index, Command = operation.Command, Status = "completed" });
                return false;
            }

            bool unknown = IsUnknownWorkflowOutcome(response.ErrorCategory);
            unknownOutcomeSeen |= unknown;
            failedIndex ??= index;
            firstFailure ??= response;
            steps.Add(new WorkflowStepReceipt
            {
                Index = index,
                Command = operation.Command,
                Status = unknown ? "unknown" : "failed",
                ErrorCategory = response.ErrorCategory,
            });
            return plan.StopOnError || unknown || terminalCheckpointFailure;
        }

        void CaptureRequestedVerification(IExcelBatch batch)
        {
            if (!verificationRequested)
            {
                return;
            }

            if (unknownOutcomeSeen)
            {
                verificationReceipt = WorkflowRangeVerifier.NotVerified(
                    verificationSheetName!,
                    verificationRangeAddress!,
                    "Verification was skipped because the plan outcome is unknown; do not inspect or replay this session until it is reconciled.");
                return;
            }

            try
            {
                verificationReceipt = WorkflowRangeVerifier.Read(
                    batch,
                    verificationSheetName!,
                    verificationRangeAddress!);
            }
            catch (Exception ex) when (
                ex is not TimeoutException and not OperationCanceledException &&
                !IsFatalExcelDisconnect(ex))
            {
                // A bad or unreadable caller-selected range does not make already-completed
                // mutations ambiguous. Report the verification limitation without encouraging
                // a replay of the plan.
                verificationReceipt = WorkflowRangeVerifier.NotVerified(
                    verificationSheetName!,
                    verificationRangeAddress!,
                    "The requested verification range could not be read. The plan receipt still describes the operation outcomes; do not replay completed mutations solely because verification failed.");
            }
        }

        async Task RunSequentialPlanAsync()
        {
            for (int index = 0; index < plan.Operations.Count; index++)
            {
                var operation = plan.Operations[index];
                attempted++;
                bool isMutation = mutationSteps[index];
                bool requestSharedCheckpoint = isMutation && sharedCheckpointRequested && !checkpointAttempted;
                checkpointAttempted |= requestSharedCheckpoint;
                var response = await ProcessAsync(CreateOperationRequest(operation, requestSharedCheckpoint)).ConfigureAwait(false);
                if (RecordOperationResponse(index, operation, isMutation, requestSharedCheckpoint, response))
                {
                    break;
                }
            }
        }

        void RunFastPlan(IExcelBatch batch, ExclusiveSessionLease lease)
        {
            batch.Execute((_, _) =>
            {
                // AsyncLocal ownership is intentionally re-established on the channel's
                // dedicated STA thread; ExecutionContext is not assumed to flow there.
                using var staOwnerScope = lease.EnterOwnerScope();
                for (int index = 0; index < plan.Operations.Count; index++)
                {
                    var operation = plan.Operations[index];
                    attempted++;
                    var response = DispatchWorkflowFastOperation(operation, batch, request.SessionId);
                    if (RecordOperationResponse(index, operation, mutationSteps[index], false, response))
                    {
                        break;
                    }
                }

                CaptureRequestedVerification(batch);

                return 0;
            });
        }

        try
        {
            using var lease = await _sessionManager.AcquireExclusiveOperationAsync(request.SessionId).ConfigureAwait(false);
            using var ownerScope = lease.EnterOwnerScope();
            var batch = _sessionManager.GetSession(request.SessionId)
                ?? throw new InvalidOperationException($"Session '{request.SessionId}' was unavailable after acquiring its workflow lease.");
            var dispatchDiagnostics = batch as IExcelBatchDispatchDiagnostics;
            long dispatchesBefore = dispatchDiagnostics?.StaDispatchCount ?? 0;
            try
            {
                if (fastModeUsed)
                {
                    RunFastPlan(batch, lease);
                }
                else
                {
                    await RunSequentialPlanAsync().ConfigureAwait(false);
                    CaptureRequestedVerification(batch);
                }
            }
            finally
            {
                if (dispatchDiagnostics is not null)
                {
                    staDispatchCount = dispatchDiagnostics.StaDispatchCount - dispatchesBefore;
                }
            }

            if (fastModeUsed && firstFailure is not null && IsUnknownWorkflowOutcome(firstFailure.ErrorCategory))
            {
                if (string.Equals(firstFailure.ErrorCategory, "ExcelProcessDied", StringComparison.Ordinal))
                {
                    RecordAndCleanupDeadSession(request.SessionId);
                }
                else
                {
                    _ = _safetyCoordinator.RecordSessionInterruption(
                        request.SessionId,
                        "abortedUnknown",
                        firstFailure.ErrorCategory ?? "UnknownOutcome");
                    _ = CloseInterruptedSession(request.SessionId);
                }
            }
        }
        catch (Exception ex)
        {
            if (attempted > 0)
            {
                if (IsFatalExcelDisconnect(ex))
                {
                    RecordAndCleanupDeadSession(request.SessionId);
                }
                else
                {
                    string category = ex switch
                    {
                        TimeoutException => "Timeout",
                        OperationCanceledException => "Cancelled",
                        _ => "UnknownOutcome",
                    };
                    _ = _safetyCoordinator.RecordSessionInterruption(request.SessionId, "abortedUnknown", category);
                    _ = CloseInterruptedSession(request.SessionId);
                }
            }

            if (attempted == 0)
            {
                firstFailure = new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "PlanNotExecuted",
                    ErrorMessage = ex.Message,
                };
            }
            else
            {
                int interruptedIndex = Math.Max(0, attempted - 1);
                failedIndex ??= interruptedIndex;
                int existingStep = steps.FindIndex(step => step.Index == interruptedIndex);
                if (existingStep >= 0)
                {
                    if (steps[existingStep].Status == "completed")
                    {
                        completed = Math.Max(0, completed - 1);
                        if (mutationSteps[interruptedIndex])
                        {
                            completedMutations = Math.Max(0, completedMutations - 1);
                        }
                    }
                    steps[existingStep] = new WorkflowStepReceipt
                    {
                        Index = interruptedIndex,
                        Command = plan.Operations[interruptedIndex].Command,
                        Status = "unknown",
                        ErrorCategory = "UnknownOutcome",
                    };
                }
                else
                {
                    steps.Add(new WorkflowStepReceipt
                    {
                        Index = interruptedIndex,
                        Command = plan.Operations[interruptedIndex].Command,
                        Status = "unknown",
                        ErrorCategory = "UnknownOutcome",
                    });
                }

                firstFailure ??= new ServiceResponse { ErrorCategory = "UnknownOutcome", ErrorMessage = ex.Message };
                unknownOutcomeSeen = true;
            }
        }
        finally
        {
            requiredCheckpointScope?.Dispose();
        }

        if (verificationRequested && verificationReceipt is null)
        {
            verificationReceipt = WorkflowRangeVerifier.NotVerified(
                verificationSheetName!,
                verificationRangeAddress!,
                unknownOutcomeSeen
                    ? "Verification was skipped because the plan outcome is unknown; do not inspect or replay this session until it is reconciled."
                    : "Verification was skipped because the plan did not reach its final read-back stage.");
        }

        for (int index = steps.Count; index < plan.Operations.Count; index++)
        {
            steps.Add(new WorkflowStepReceipt { Index = index, Command = plan.Operations[index].Command, Status = "notStarted" });
        }

        var receipt = new WorkflowPlanReceipt
        {
            PlanId = planId,
            Outcome = firstFailure is null ? WorkflowPlanOutcome.Completed :
                (unknownOutcomeSeen || IsUnknownWorkflowOutcome(firstFailure.ErrorCategory)
                    ? WorkflowPlanOutcome.Unknown : WorkflowPlanOutcome.Failed),
            OperationCount = plan.Operations.Count,
            AttemptedCount = attempted,
            CompletedCount = completed,
            FailedIndex = failedIndex,
            Checkpoint = checkpointReceipt,
            Steps = steps,
            ExecutionMode = executionMode,
            FastModeRequested = plan.FastMode,
            FastModeUsed = fastModeUsed,
            FastModeFallbackReason = fastModeFallbackReason,
            StaDispatchCount = staDispatchCount,
            Verification = verificationReceipt,
        };
        bool unknownOutcome = receipt.Outcome == WorkflowPlanOutcome.Unknown;
        string? aggregateErrorCategory = firstFailure switch
        {
            null => null,
            _ when unknownOutcome => "UnknownOutcome",
            _ when attempted == 0 => "PlanNotExecuted",
            _ when knownCheckpointFailure && completedMutations == 0 => "CheckpointFailed",
            _ => "PlanFailed",
        };
        return new ServiceResponse
        {
            Success = firstFailure is null,
            Command = request.Command,
            SessionId = request.SessionId,
            ErrorCategory = aggregateErrorCategory,
            ErrorMessage = firstFailure?.ErrorMessage,
            Result = ServiceProtocol.Serialize(receipt),
        };
    }

    /// <summary>
    /// Dispatches only commands admitted by <see cref="WorkflowFastPathPolicy"/>.
    /// The caller already owns the session lease and the batch's queued STA work item,
    /// so this deliberately bypasses per-step session admission and safety orchestration.
    /// Fast-path selection requires the complete safety workflow to be off.
    /// </summary>
    private ServiceResponse DispatchWorkflowFastOperation(
        ServiceBatchOperation operation,
        IExcelBatch batch,
        string sessionId)
    {
        var parts = operation.Command.Split('.', 2);
        string action = parts.Length == 2 ? parts[1] : string.Empty;
        string? args = GetOperationArgsJson(operation.Args);

        try
        {
            switch (parts[0])
            {
                case "range" when ServiceRegistry.Range.TryParseAction(action, out var rangeAction):
                    return WrapResult(ServiceRegistry.Range.DispatchToCore(_rangeCommands, rangeAction, batch, args));
                case "rangeedit" when ServiceRegistry.RangeEdit.TryParseAction(action, out var rangeEditAction):
                    return WrapResult(ServiceRegistry.RangeEdit.DispatchToCore(_rangeCommands, rangeEditAction, batch, args));
                case "rangeformat" when ServiceRegistry.RangeFormat.TryParseAction(action, out var rangeFormatAction):
                    return WrapResult(ServiceRegistry.RangeFormat.DispatchToCore(_rangeCommands, rangeFormatAction, batch, args));
                case "rangelink" when ServiceRegistry.RangeLink.TryParseAction(action, out var rangeLinkAction):
                    return WrapResult(ServiceRegistry.RangeLink.DispatchToCore(_rangeCommands, rangeLinkAction, batch, args));
                case "sheet" when ServiceRegistry.Sheet.TryParseAction(action, out var sheetAction):
                    return WrapResult(ServiceRegistry.Sheet.DispatchToCore(_sheetCommands, sheetAction, batch, args));
                case "sheetstyle" when ServiceRegistry.SheetStyle.TryParseAction(action, out var sheetStyleAction):
                    return WrapResult(ServiceRegistry.SheetStyle.DispatchToCore(_sheetCommands, sheetStyleAction, batch, args));
                case "reportformat" when ServiceRegistry.ReportFormat.TryParseAction(action, out var reportFormatAction):
                    return WrapResult(ServiceRegistry.ReportFormat.DispatchToCore(_reportFormatCommands, reportFormatAction, batch, args));
                default:
                    return new ServiceResponse
                    {
                        Success = false,
                        Command = operation.Command,
                        SessionId = sessionId,
                        ErrorCategory = "PlanSafetyConflict",
                        ErrorMessage = $"Command '{operation.Command}' is not available in the one-STA workflow executor."
                    };
            }
        }
        catch (OperationCanceledException ex)
        {
            return new ServiceResponse
            {
                Success = false,
                Command = operation.Command,
                SessionId = sessionId,
                ErrorCategory = "Cancelled",
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name,
            };
        }
        catch (Exception ex) when (IsFatalExcelDisconnect(ex))
        {
            return CreateExcelDisconnectedResponse(
                sessionId,
                ex,
                $"Excel disconnected during fast workflow command '{operation.Command}'. The session will be cleaned up; reopen the file before retrying.");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex, operation.Command, sessionId);
        }
    }

    private static bool IsUnknownWorkflowOutcome(string? category) => category is
        "Timeout" or "Cancelled" or "Canceled" or "ExcelProcessDied" or "UnknownOutcome" or
        "AbortedUnknown" or "JournalPersistenceFailed" or "IdempotencyUnknownOutcome" or
        "IdempotencyInProgress" or "SessionInterrupted" or "ServerShutdown";

    private static WorkflowCheckpointReceipt? TryGetCheckpointReceipt(string? responseResult)
    {
        if (string.IsNullOrWhiteSpace(responseResult))
        {
            return null;
        }

        try
        {
            using var document = JsonDocument.Parse(responseResult);
            return document.RootElement.TryGetProperty("checkpoint", out var checkpoint)
                ? TryParseCheckpointReceipt(checkpoint)
                : null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private WorkflowCheckpointReceipt? TryGetLatestCheckpointReceipt(
        string sessionId,
        DateTime planStartedAtUtc,
        string command)
    {
        try
        {
            using var document = JsonDocument.Parse(_safetyCoordinator.GetJournal(sessionId).Result ?? "{}");
            if (!document.RootElement.TryGetProperty("operations", out var operations) || operations.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            foreach (var operation in operations.EnumerateArray().Reverse())
            {
                if (!operation.TryGetProperty("createdAtUtc", out var createdAt) ||
                    !DateTime.TryParse(createdAt.GetString(), out var createdAtUtc) ||
                    createdAtUtc.ToUniversalTime() < planStartedAtUtc.ToUniversalTime())
                {
                    continue;
                }
                if (!operation.TryGetProperty("command", out var journalCommand) ||
                    !string.Equals(journalCommand.GetString(), command, StringComparison.Ordinal))
                {
                    continue;
                }
                if (!operation.TryGetProperty("checkpoint", out var checkpoint) || checkpoint.ValueKind != JsonValueKind.Object)
                {
                    continue;
                }

                if (TryParseCheckpointReceipt(checkpoint) is { } receipt)
                {
                    return receipt;
                }
            }
        }
        catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException)
        {
            // Receipt evidence is best effort; the operation journal remains authoritative.
        }

        return null;
    }

    private static WorkflowCheckpointReceipt? TryParseCheckpointReceipt(JsonElement checkpoint)
    {
        if (checkpoint.ValueKind != JsonValueKind.Object ||
            !checkpoint.TryGetProperty("recoveryId", out var recoveryIdElement) ||
            !checkpoint.TryGetProperty("relativePath", out var relativePathElement) ||
            !checkpoint.TryGetProperty("sha256", out var sha256Element) ||
            !checkpoint.TryGetProperty("size", out var sizeElement) ||
            sizeElement.ValueKind != JsonValueKind.Number ||
            !sizeElement.TryGetInt64(out long size))
        {
            return null;
        }

        string? recoveryId = recoveryIdElement.GetString();
        string? relativePath = relativePathElement.GetString();
        string? sha256 = sha256Element.GetString();
        if (string.IsNullOrWhiteSpace(recoveryId) ||
            string.IsNullOrWhiteSpace(relativePath) ||
            sha256?.Length != 64 ||
            !sha256.All(Uri.IsHexDigit) ||
            size <= 0)
        {
            return null;
        }

        return new WorkflowCheckpointReceipt
        {
            RecoveryId = recoveryId,
            RelativePath = relativePath,
            Sha256 = sha256,
            Size = size,
        };
    }

    private void CleanupOwnedWorkflowSession(string sessionId)
    {
        try
        {
            _sessionManager.CloseSession(sessionId, save: false, force: true);
        }
        catch (Exception cleanupEx)
        {
            System.Diagnostics.Debug.WriteLine($"Workflow session cleanup failed for {sessionId}: {cleanupEx.Message}");
        }
        finally
        {
            _safetyCoordinator.RemoveSession(sessionId);
            _idempotencyCoordinator.RemoveSession(sessionId);
        }
    }

    private ServiceResponse HandleSessionCommand(string action, ServiceRequest request)
    {
        return action switch
        {
            "create" => HandleSessionCreate(request),
            "open" => HandleSessionOpen(request),
            "close" => HandleSessionClose(request),
            "save" => HandleSessionSave(request),
            "list" => HandleSessionList(),
            "preflight" => HandleSessionPreflight(request),
            "configure-safety" => HandleSessionConfigureSafety(request),
            "journal" => HandleSessionJournal(request),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown session action: {action}" }
        };
    }

    private async Task<ServiceResponse> HandleSessionBatchAsync(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "sessionId is required",
            };
        }

        ServiceBatchRequest? batchRequest;
        try
        {
            batchRequest = ServiceProtocol.Deserialize<ServiceBatchRequest>(request.Args ?? "{}");
        }
        catch (JsonException ex)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Invalid session.batch arguments: {ex.Message}",
            };
        }

        if (batchRequest?.Operations is not { Count: > 0 })
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "session.batch requires at least one operation",
            };
        }

        if (batchRequest.Operations.Count > MaximumBatchOperations)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "BatchTooLarge",
                ErrorMessage = $"session.batch accepts at most {MaximumBatchOperations} operations",
            };
        }

        for (int index = 0; index < batchRequest.Operations.Count; index++)
        {
            var operation = batchRequest.Operations[index];
            string? validationError = ValidateBatchOperation(request.SessionId, operation);
            if (validationError != null)
            {
                return CreateBatchValidationFailure(index, operation.Command, validationError);
            }
        }

        var results = new List<ServiceBatchOperationResult>(batchRequest.Operations.Count);
        int? failedIndex = null;
        ServiceResponse? firstFailure = null;

        for (int index = 0; index < batchRequest.Operations.Count; index++)
        {
            var operation = batchRequest.Operations[index];
            var operationRequest = new ServiceRequest
            {
                Command = operation.Command,
                SessionId = request.SessionId,
                Args = GetOperationArgsJson(operation.Args),
                Source = request.Source,
                ReviewOnly = operation.ReviewOnly,
                ReviewId = operation.ReviewId,
                Checkpoint = operation.Checkpoint,
                IdempotencyKey = operation.IdempotencyKey,
            };

            var operationResponse = await ProcessAsync(operationRequest);
            results.Add(new ServiceBatchOperationResult
            {
                Index = index,
                Success = operationResponse.Success,
                Result = ParseEmbeddedResult(operationResponse.Result),
                ErrorMessage = operationResponse.ErrorMessage,
                ErrorCategory = operationResponse.ErrorCategory,
                ExceptionType = operationResponse.ExceptionType,
                HResult = operationResponse.HResult,
            });

            if (!operationResponse.Success)
            {
                failedIndex ??= index;
                firstFailure ??= operationResponse;
                if (batchRequest.StopOnError || IsTerminalBatchFailure(operationResponse))
                {
                    break;
                }
            }
        }

        bool completed = results.Count == batchRequest.Operations.Count;
        var batchResponse = new ServiceBatchResponse
        {
            Success = failedIndex == null,
            Completed = completed,
            FailedIndex = failedIndex,
            Results = results,
        };

        return new ServiceResponse
        {
            Success = batchResponse.Success,
            ErrorMessage = firstFailure?.ErrorMessage,
            ErrorCategory = firstFailure?.ErrorCategory,
            ExceptionType = firstFailure?.ExceptionType,
            HResult = firstFailure?.HResult,
            Result = ServiceProtocol.Serialize(batchResponse),
        };
    }

    private static string? ValidateBatchOperation(string sessionId, ServiceBatchOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Command))
        {
            return "Every batch operation requires a command";
        }

        if (!string.IsNullOrWhiteSpace(operation.SessionId) &&
            !string.Equals(operation.SessionId, sessionId, StringComparison.Ordinal))
        {
            return "A batch operation cannot target a different session";
        }

        var parts = operation.Command.Split('.', 2);
        if (parts.Length != 2 || string.IsNullOrWhiteSpace(parts[1]))
        {
            return $"Invalid batch command: {operation.Command}";
        }

        bool sessionScoped = parts[0] switch
        {
            "sheet" => !ServiceRegistry.Sheet.TryParseAction(parts[1], out var action) ||
                ServiceRegistry.Sheet.RequiresSessionForAction(action),
            "sheetstyle" or
            "range" or "rangeedit" or "rangeformat" or "rangelink" or
            "table" or "tablecolumn" or
            "powerquery" or
            "pivottable" or "pivottablefield" or "pivottablecalc" or
            "chart" or "chartconfig" or
            "connection" or "calculation" or "namedrange" or
            "conditionalformat" or "vba" or
            "datamodel" or "datamodelrel" or
            "slicer" or "screenshot" or "window" or "pythoninexcel" => true,
            _ => false,
        };

        return sessionScoped
            ? null
            : $"Command '{operation.Command}' is not allowed inside session.batch";
    }

    private static ServiceResponse CreateBatchValidationFailure(int index, string command, string message)
    {
        var result = new ServiceBatchResponse
        {
            Success = false,
            Completed = false,
            FailedIndex = index,
            Results =
            [
                new ServiceBatchOperationResult
                {
                    Index = index,
                    Command = command,
                    Success = false,
                    ErrorMessage = message,
                    ErrorCategory = "InvalidInput",
                },
            ],
        };

        return new ServiceResponse
        {
            Success = false,
            ErrorCategory = "InvalidInput",
            ErrorMessage = message,
            Result = ServiceProtocol.Serialize(result),
        };
    }

    private static string? GetOperationArgsJson(JsonElement? args)
    {
        if (args is not { } value || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return null;
        }

        return value.ValueKind == JsonValueKind.String ? value.GetString() : value.GetRawText();
    }

    private static JsonElement? ParseEmbeddedResult(string? result)
    {
        if (string.IsNullOrWhiteSpace(result))
        {
            return null;
        }

        try
        {
            using var document = JsonDocument.Parse(result);
            return document.RootElement.Clone();
        }
        catch (JsonException)
        {
            return JsonSerializer.SerializeToElement(result, ServiceProtocol.JsonOptions);
        }
    }

    private static bool IsTerminalBatchFailure(ServiceResponse response) =>
        response.ErrorCategory is "Timeout" or "Cancelled" or "ExcelProcessDied" or
            "IdempotencyUnknownOutcome" or "UnknownOutcome" or "AbortedUnknown" or "JournalPersistenceFailed";

    private ServiceResponse HandleSessionCreate(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        var validation = ExcelFileValidator.Inspect(args.FilePath);
        if (!validation.IsWithinPathLimit)
        {
            return new ServiceResponse { Success = false, ErrorMessage = validation.Message };
        }

        if (!validation.IsWithinCreatePathLimit)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "PathTooLong",
                ErrorMessage = $"File path exceeds Excel's practical SaveAs limit of {ExcelFileValidator.MaximumCreatePathLength} characters."
            };
        }

        var fullPath = validation.FilePath;

        if (File.Exists(fullPath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"File already exists: {fullPath}. Use session open to open an existing workbook."
            };
        }

        if (!validation.IsSupportedExtension)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Invalid file extension '{validation.Extension}'. session create supports .xlsx and .xlsm only."
            };
        }

        try
        {
            // Use the combined create+open which starts Excel only once
            TimeSpan? timeout = args.TimeoutSeconds.HasValue
                ? TimeSpan.FromSeconds(args.TimeoutSeconds.Value)
                : null;
            var sessionId = _sessionManager.CreateSessionForNewFile(fullPath, show: args.Show, operationTimeout: timeout, origin: SessionOrigin.CLI);
            _knownSessionIds.TryAdd(sessionId, 0);

            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId, filePath = fullPath }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex);
        }
    }

    private ServiceResponse HandleSessionOpen(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        try
        {
            TimeSpan? timeout = args.TimeoutSeconds.HasValue
                ? TimeSpan.FromSeconds(args.TimeoutSeconds.Value)
                : null;
            var sessionId = _sessionManager.CreateSession(args.FilePath, show: args.Show, operationTimeout: timeout, origin: SessionOrigin.CLI);
            _knownSessionIds.TryAdd(sessionId, 0);
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId, filePath = args.FilePath }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex);
        }
    }

    private ServiceResponse HandleSessionClose(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var args = ServiceRegistry.DeserializeArgs<SessionCloseArgs>(request.Args);
        var shouldSave = args?.Save ?? false;
        bool closed;
        bool excelProcessWasDead;
        try
        {
            closed = _sessionManager.CloseSession(
                request.SessionId,
                save: shouldSave,
                force: false,
                out excelProcessWasDead);
        }
        catch (Exception ex) when (IsFatalExcelDisconnect(ex))
        {
            RecordAndCleanupDeadSession(request.SessionId);
            return CreateExcelDisconnectedResponse(request.SessionId, ex, shouldSave
                ? "Excel disconnected while saving before close. Session has been cleaned up; reopen the workbook and verify whether the save completed."
                : "Excel disconnected while closing. Session has been cleaned up; reopen the workbook with a new session.");
        }

        if (excelProcessWasDead)
        {
            return new ServiceResponse
            {
                Success = false,
                SessionId = request.SessionId,
                ErrorCategory = "ExcelProcessDied",
                ErrorMessage = shouldSave
                    ? "Excel had already stopped before close, so changes could not be saved. Recovery evidence was recorded; reopen the workbook and verify its contents."
                    : "Excel had already stopped before close. Recovery evidence was recorded and the dead session was cleaned up."
            };
        }

        if (closed)
        {
            _safetyCoordinator.RemoveSession(request.SessionId);
            _idempotencyCoordinator.RemoveSession(request.SessionId);
            return new ServiceResponse { Success = true };
        }

        if (_knownSessionIds.ContainsKey(request.SessionId))
        {
            _safetyCoordinator.RemoveSession(request.SessionId);
            _idempotencyCoordinator.RemoveSession(request.SessionId);
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(
                    new { success = true, sessionId = request.SessionId, message = "Session already closed." },
                    ServiceProtocol.JsonOptions)
            };
        }

        return new ServiceResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
    }

    private ServiceResponse HandleSessionSave(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var sessionError = TryBeginUsableSession(request.SessionId, out var batch);
        if (sessionError != null)
        {
            return sessionError;
        }

        try
        {
            batch!.Save();
            return new ServiceResponse { Success = true };
        }
        catch (Exception ex) when (IsFatalExcelDisconnect(ex))
        {
            RecordAndCleanupDeadSession(request.SessionId);
            return CreateExcelDisconnectedResponse(request.SessionId, ex,
                "Excel disconnected while saving. Session has been cleaned up; reopen the workbook and verify whether the save completed.");
        }
        finally
        {
            _sessionManager.EndOperation(request.SessionId);
        }
    }

    private ServiceResponse HandleSessionList()
    {
        var sessions = _sessionManager.GetActiveSessions()
            .Select(s => new
            {
                sessionId = s.SessionId,
                filePath = s.FilePath,
                isExcelVisible = _sessionManager.IsExcelVisible(s.SessionId),
                activeOperations = _sessionManager.GetActiveOperationCount(s.SessionId),
                canClose = _sessionManager.GetActiveOperationCount(s.SessionId) == 0
            })
            .ToList();

        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new { success = true, sessions, count = sessions.Count }, ServiceProtocol.JsonOptions)
        };
    }

    private ServiceResponse HandleSessionConfigureSafety(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        if (!_sessionManager.IsSessionAlive(request.SessionId) ||
            !_sessionManager.TryGetFilePath(request.SessionId, out var workbookPath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Session '{request.SessionId}' not found or Excel is no longer running."
            };
        }

        var response = _safetyCoordinator.Configure(request.SessionId, workbookPath, request.Args);
        if (response.Success)
        {
            _sessionManager.SetAutoSaveOnDispose(
                request.SessionId,
                autoSave: !_safetyCoordinator.ShouldDiscardOnAbnormalShutdown(request.SessionId));
        }

        return response;
    }

    private ServiceResponse HandleSessionJournal(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        return _safetyCoordinator.GetJournal(request.SessionId);
    }

    private ServiceResponse HandleRecoveryCommand(string action, ServiceRequest request)
    {
        if (action == "list")
        {
            return _safetyCoordinator.ListRecoveries();
        }

        if (action != "recover")
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown recovery action: {action}" };
        }

        var args = ServiceRegistry.DeserializeArgs<RecoveryArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args?.RecoveryId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "recoveryId is required" };
        }

        if (!_safetyCoordinator.TryResolveRecovery(args.RecoveryId, out var checkpointPath, out var operationId) ||
            string.IsNullOrWhiteSpace(checkpointPath) || string.IsNullOrWhiteSpace(operationId))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "RecoveryUnavailable",
                ErrorMessage = $"Recovery '{args.RecoveryId}' is unavailable or failed integrity validation."
            };
        }

        try
        {
            TimeSpan? timeout = args.TimeoutSeconds.HasValue
                ? TimeSpan.FromSeconds(args.TimeoutSeconds.Value)
                : null;
            var origin = request.Source?.StartsWith("mcp", StringComparison.OrdinalIgnoreCase) == true
                ? SessionOrigin.MCP
                : SessionOrigin.CLI;
            var sessionId = _sessionManager.CreateSession(
                checkpointPath,
                show: args.Show,
                operationTimeout: timeout,
                origin: origin);
            _knownSessionIds.TryAdd(sessionId, 0);
            _safetyCoordinator.RecordRecovered(operationId);
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new
                {
                    success = true,
                    sessionId,
                    recoveryId = args.RecoveryId,
                    originalWorkbookOverwritten = false
                }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex, request.Command, request.SessionId);
        }
    }

    private ServiceResponse HandleSessionPreflight(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var sessionError = TryBeginUsableSession(request.SessionId, out var batch);
        if (sessionError != null)
        {
            return sessionError;
        }

        try
        {
            var result = CapabilityPreflightCommands.Collect(batch!, request.SessionId);
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions)
            };
        }
        catch (ExcelOperationNotStartedTimeoutException ex)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "TimeoutBeforeExecution",
                ErrorMessage = $"Excel operation timed out before execution. The session remains open and retrying is safe: {ex.Message}",
                ExceptionType = ex.GetType().Name
            };
        }
        catch (ExcelOperationNotStartedCanceledException ex)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "CancelledBeforeExecution",
                ErrorMessage = $"Excel operation was cancelled before execution. The session remains open and retrying is safe: {ex.Message}",
                ExceptionType = ex.GetType().Name
            };
        }
        catch (TimeoutException ex)
        {
            _safetyCoordinator.RecordSessionInterruption(request.SessionId, "abortedUnknown", "Timeout");
            bool closed = CloseInterruptedSession(request.SessionId);
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Timeout",
                ErrorMessage = closed
                    ? $"Excel capability preflight timed out and the session has been closed: {ex.Message}"
                    : $"Excel capability preflight timed out, but session closure could not be confirmed: {ex.Message} Do not reuse this session; close Excel and reopen the file.",
                ExceptionType = ex.GetType().Name
            };
        }
        catch (OperationCanceledException)
        {
            _safetyCoordinator.RecordSessionInterruption(request.SessionId, "abortedUnknown", "Cancelled");
            bool closed = CloseInterruptedSession(request.SessionId);
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Cancelled",
                ErrorMessage = closed
                    ? "Excel capability preflight was cancelled and the session has been closed."
                    : "Excel capability preflight was cancelled, but session closure could not be confirmed. Do not reuse this session; close Excel and reopen the file.",
                ExceptionType = nameof(OperationCanceledException)
            };
        }
        catch (Exception ex) when (IsFatalExcelDisconnect(ex))
        {
            RecordAndCleanupDeadSession(request.SessionId);
            return CreateExcelDisconnectedResponse(
                request.SessionId,
                ex,
                $"Excel process for session '{request.SessionId}' disconnected during capability preflight. Session has been cleaned up. Please reopen the file with a new session.");
        }
        catch (Exception ex)
        {
            if (batch is not null && !batch.IsExcelProcessAlive())
            {
                RecordAndCleanupDeadSession(request.SessionId);
                return CreateExcelDisconnectedResponse(
                    request.SessionId,
                    ex,
                    $"Excel process for session '{request.SessionId}' died during capability preflight. Session has been cleaned up. Please reopen the file with a new session.");
            }

            return CreateErrorResponse(ex, request.Command, request.SessionId);
        }
        finally
        {
            _sessionManager.EndOperation(request.SessionId);
        }
    }



    // === GENERATED DISPATCH ===

    // All command routing uses ServiceRegistry.*.DispatchToCore() generated methods.

    // See ServiceRegistry.*.Dispatch.g.cs for the generated code.



    private delegate bool TryParseDelegate<TAction>(string action, out TAction result);



    private static ServiceResponse WrapResult(string? dispatchResult)

    {

        return dispatchResult == null

            ? new ServiceResponse { Success = true }

            : new ServiceResponse { Success = true, Result = dispatchResult };

    }



    private async Task<ServiceResponse> DispatchSimpleAsync<TAction>(

        string actionString, ServiceRequest request,

        TryParseDelegate<TAction> tryParse,

        Func<TAction, IExcelBatch, string?> dispatch) where TAction : struct

    {

        if (!tryParse(actionString, out var action))

            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };



        return await WithSessionAsync(request, batch => WrapResult(dispatch(action, batch)));

    }

    /// <summary>
    /// Dispatches a session-less command (no Excel batch required).
    /// Used for [NoSession] categories like diag.
    /// </summary>
    private ServiceResponse DispatchSessionless(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.Diag.TryParseAction(actionString, out var action))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown action: {actionString}" };

        return WrapResult(ServiceRegistry.Diag.DispatchToCore(_diagCommands, action, request.Args));
    }

    private async Task<ServiceResponse> DispatchSheetAsync(string actionString, ServiceRequest request)

    {

        if (ServiceRegistry.Sheet.TryParseAction(actionString, out var sheetAction))

        {

            // Atomic actions are self-contained and do not require a session.
            if (!ServiceRegistry.Sheet.RequiresSessionForAction(sheetAction))

            {

                if (request.ReviewOnly || !string.IsNullOrWhiteSpace(request.ReviewId) || request.Checkpoint)

                {

                    return new ServiceResponse

                    {

                        Success = false,

                        ErrorCategory = "SafetyWorkflowUnavailable",

                        ErrorMessage = "Atomic cross-file worksheet actions do not yet support review or checkpoints; neither workbook was changed. Run without safety options only after reviewing the source and target files directly."

                    };

                }

                try

                {

                    return WrapResult(ServiceRegistry.Sheet.DispatchToCore(

                        _sheetCommands, sheetAction, null!, request.Args));

                }

                catch (Exception ex)

                {

                    return CreateErrorResponse(ex);

                }

            }



            return await WithSessionAsync(request, batch =>

                WrapResult(ServiceRegistry.Sheet.DispatchToCore(_sheetCommands, sheetAction, batch, request.Args)));

        }



        if (ServiceRegistry.SheetStyle.TryParseAction(actionString, out var styleAction))

        {

            return await WithSessionAsync(request, batch =>

                WrapResult(ServiceRegistry.SheetStyle.DispatchToCore(_sheetCommands, styleAction, batch, request.Args)));

        }



        return new ServiceResponse { Success = false, ErrorMessage = $"Unknown sheet action: {actionString}" };

    }



    private async Task<ServiceResponse> DispatchRangeAsync(string actionString, ServiceRequest request)

    {

        return await WithSessionAsync(request, batch =>

        {

            if (ServiceRegistry.Range.TryParseAction(actionString, out var ra))

                return WrapResult(ServiceRegistry.Range.DispatchToCore(_rangeCommands, ra, batch, request.Args));

            if (ServiceRegistry.RangeEdit.TryParseAction(actionString, out var rea))

                return WrapResult(ServiceRegistry.RangeEdit.DispatchToCore(_rangeCommands, rea, batch, request.Args));

            if (ServiceRegistry.RangeFormat.TryParseAction(actionString, out var rfa))

                return WrapResult(ServiceRegistry.RangeFormat.DispatchToCore(_rangeCommands, rfa, batch, request.Args));

            if (ServiceRegistry.RangeLink.TryParseAction(actionString, out var rla))

                return WrapResult(ServiceRegistry.RangeLink.DispatchToCore(_rangeCommands, rla, batch, request.Args));

            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown range action: {actionString}" };

        });

    }



    private async Task<ServiceResponse> DispatchTableAsync(string actionString, ServiceRequest request)

    {

        return await WithSessionAsync(request, batch =>

        {

            if (ServiceRegistry.Table.TryParseAction(actionString, out var ta))

                return WrapResult(ServiceRegistry.Table.DispatchToCore(_tableCommands, ta, batch, request.Args));

            if (ServiceRegistry.TableColumn.TryParseAction(actionString, out var tca))

                return WrapResult(ServiceRegistry.TableColumn.DispatchToCore(_tableCommands, tca, batch, request.Args));

            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown table action: {actionString}" };

        });

    }

    private async Task<ServiceResponse> DispatchWindowAsync(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.Window.TryParseAction(actionString, out var windowAction))
            return new ServiceResponse { Success = false, ErrorMessage = $"Unknown window action: {actionString}" };

        return await WithSessionAsync(request, batch =>
        {
            var result = WrapResult(ServiceRegistry.Window.DispatchToCore(_windowCommands, windowAction, batch, request.Args));

            // Update SessionManager visibility flag when show/hide commands succeed
            if (result.Success && !string.IsNullOrWhiteSpace(request.SessionId))
            {
                if (windowAction is WindowAction.Show or WindowAction.Arrange or WindowAction.SetState or WindowAction.SetPosition)
                {
                    _sessionManager.SetExcelVisible(request.SessionId, true);
                }
                else if (windowAction is WindowAction.Hide)
                {
                    _sessionManager.SetExcelVisible(request.SessionId, false);
                }
            }

            return result;
        });
    }


    private Task<ServiceResponse> WithSessionAsync(ServiceRequest request, Func<IExcelBatch, ServiceResponse> action)
    {
        var sessionId = request.SessionId;
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return Task.FromResult(new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" });
        }

        var sessionError = TryBeginUsableSession(sessionId, out var batch);
        if (sessionError != null)
        {
            return Task.FromResult(sessionError);
        }

        try
        {
            var response = _safetyCoordinator.Execute(request, batch!, () => action(batch!));
            return Task.FromResult(response);
        }
        catch (ExcelOperationNotStartedTimeoutException ex)
        {
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "TimeoutBeforeExecution",
                ErrorMessage = $"Excel operation timed out before execution. The session remains open and retrying is safe: {ex.Message}",
                ExceptionType = ex.GetType().Name
            });
        }
        catch (ExcelOperationNotStartedCanceledException ex)
        {
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "CancelledBeforeExecution",
                ErrorMessage = $"Excel operation was cancelled before execution. The session remains open and retrying is safe: {ex.Message}",
                ExceptionType = ex.GetType().Name
            });
        }
        catch (TimeoutException ex)
        {
            _safetyCoordinator.RecordSessionInterruption(sessionId, "abortedUnknown", "Timeout");
            // Operation timed out — Excel COM call is hung (IDispatch.Invoke stuck).
            // Force-close the session to trigger the force-kill path in ExcelBatch.Dispose(),
            // which will kill the hung Excel process and release the STA thread.
            bool closed = CloseInterruptedSession(sessionId);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Timeout",
                ErrorMessage = closed
                    ? $"Excel operation timed out and the session has been closed: {ex.Message} Please reopen the file with a new session."
                    : $"Excel operation timed out, but session closure could not be confirmed: {ex.Message} Do not reuse this session; close Excel and reopen the file.",
                ExceptionType = ex.GetType().Name
            });
        }
        catch (OperationCanceledException)
        {
            _safetyCoordinator.RecordSessionInterruption(sessionId, "abortedUnknown", "Cancelled");
            // Caller cancelled (e.g., VS Code cancelled the tool call) while a COM operation
            // may still be running on the STA thread. ExcelBatch.Execute sets _operationTimedOut
            // on cancellation, but nobody calls Dispose() — the session stays alive with a
            // stuck STA thread, and all subsequent requests queue up and hang.
            // Force-close the session to kill the hung Excel process and release the STA thread.
            bool closed = CloseInterruptedSession(sessionId);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Cancelled",
                ErrorMessage = closed
                    ? "Operation was cancelled and the session has been closed. The Excel COM thread may have been unresponsive. Please reopen the file with a new session."
                    : "Operation was cancelled, but session closure could not be confirmed. Do not reuse this session; close Excel and reopen the file.",
                ExceptionType = nameof(OperationCanceledException)
            });
        }
        catch (COMException ex) when (
            ex.HResult == ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE ||
            ex.HResult == ResiliencePipelines.RPC_E_CALL_FAILED ||
            ex.HResult == ResiliencePipelines.RPC_E_DISCONNECTED)
        {
            RecordAndCleanupDeadSession(sessionId);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "ExcelProcessDied",
                ErrorMessage = $"Excel process for session '{sessionId}' has died (the application may have been closed or crashed). " +
                               "Session has been cleaned up. Please reopen the file with a new session.",
                ExceptionType = ex.GetType().Name,
                HResult = $"0x{ex.HResult:X8}"
            });
        }
        catch (InvalidOperationException ex) when (
            ex.Message.Contains("no longer running", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("process", StringComparison.OrdinalIgnoreCase))
        {
            RecordAndCleanupDeadSession(sessionId);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "ExcelProcessDied",
                ErrorMessage = $"Excel process for session '{sessionId}' is no longer running. " +
                               "Session has been cleaned up. Please reopen the file with a new session.",
                ExceptionType = ex.GetType().Name
            });
        }
        catch (Exception ex)
        {
            if (IsFatalExcelDisconnect(ex))
            {
                RecordAndCleanupDeadSession(sessionId);
                return Task.FromResult(CreateExcelDisconnectedResponse(sessionId, ex,
                    $"Excel process for session '{sessionId}' disconnected during the operation. Session has been cleaned up. Please reopen the file with a new session."));
            }

            // Check if Excel died with a non-COM exception — clean up dead session
            if (batch != null && !batch.IsExcelProcessAlive())
            {
                RecordAndCleanupDeadSession(sessionId);
            }

            return Task.FromResult(CreateErrorResponse(ex));
        }
        finally
        {
            _sessionManager.EndOperation(sessionId);
        }
    }

    private void CleanupDeadSession(string sessionId)
    {
        try
        {
            _sessionManager.CloseSession(sessionId, save: false, force: true);
        }
        catch (Exception cleanupEx)
        {
            System.Diagnostics.Debug.WriteLine($"Session cleanup failed for {sessionId}: {cleanupEx.Message}");
        }
        finally
        {
            _safetyCoordinator.RemoveSession(sessionId);
        }
    }

    private void RecordAndCleanupDeadSession(string sessionId)
    {
        _ = _safetyCoordinator.RecordSessionInterruption(
            sessionId,
            "excelProcessDied",
            "ExcelProcessDied");
        CleanupDeadSession(sessionId);
    }

    private void HandleDeadSessionCleanupStarting(string sessionId)
    {
        _ = _safetyCoordinator.RecordSessionInterruption(
            sessionId,
            "excelProcessDied",
            "ExcelProcessDied");
        _safetyCoordinator.RemoveSession(sessionId);
        _idempotencyCoordinator.RemoveSession(sessionId);
    }

    private static ServiceResponse CreateExcelDisconnectedResponse(string sessionId, Exception ex, string message)
    {
        return new ServiceResponse
        {
            Success = false,
            SessionId = sessionId,
            ErrorCategory = "ExcelProcessDied",
            ErrorMessage = message,
            ExceptionType = ex.GetType().Name,
            HResult = TryGetFatalComHResult(ex) is { } hresult ? $"0x{hresult:X8}" : null,
            InnerError = ex.InnerException?.Message
        };
    }

    private static bool IsFatalExcelDisconnect(Exception ex) => TryGetFatalComHResult(ex).HasValue;

    private static int? TryGetFatalComHResult(Exception ex)
    {
        for (var current = ex; current != null; current = current.InnerException!)
        {
            if (current is COMException comEx &&
                (IsFatalComHResult(comEx.HResult) || IsFatalComHResult(comEx.ErrorCode)))
            {
                return IsFatalComHResult(comEx.HResult) ? comEx.HResult : comEx.ErrorCode;
            }

            if (current.Message.Contains("disconnected", StringComparison.OrdinalIgnoreCase))
            {
                return ResiliencePipelines.RPC_E_DISCONNECTED;
            }

            if (current.Message.Contains("RPC server is unavailable", StringComparison.OrdinalIgnoreCase))
            {
                return ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE;
            }
        }

        return null;
    }

    private static bool IsFatalComHResult(int hresult) =>
        hresult == ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE ||
        hresult == ResiliencePipelines.RPC_E_CALL_FAILED ||
        hresult == ResiliencePipelines.RPC_E_DISCONNECTED;

    private static ServiceResponse AttachRequestContext(ServiceRequest request, ServiceResponse response)
    {
        if (response.Success)
        {
            return response;
        }

        var command = response.Command ?? request.Command;
        var sessionId = response.SessionId ?? request.SessionId;

        return CloneResponse(response, command, sessionId);
    }

    private static ServiceResponse CloneResponse(ServiceResponse response, string? command, string? sessionId)
    {
        return new ServiceResponse
        {
            Success = response.Success,
            Command = command,
            SessionId = sessionId,
            ErrorMessage = SensitiveDataSanitizer.Redact(response.ErrorMessage),
            ErrorCategory = response.ErrorCategory,
            ExceptionType = response.ExceptionType,
            HResult = response.HResult,
            InnerError = SensitiveDataSanitizer.Redact(response.InnerError),
            Result = response.Result
        };
    }

    private static ServiceResponse CreateErrorResponse(Exception ex, string? command = null, string? sessionId = null)
    {
        var exceptionType = ex.GetType().Name;
        string? hresult = ex is COMException comEx ? $"0x{comEx.HResult:X8}" : null;
        string? innerError = null;
        var errorCategory = ex switch
        {
            PowerQueryCommandException pqEx => pqEx.ErrorCategory,
            TimeoutException => "Timeout",
            ArgumentException => "InvalidInput",
            COMException => "ComInterop",
            _ => null
        };

        if (ex.InnerException != null)
        {
            innerError = SensitiveDataSanitizer.Redact(ex.InnerException.Message);
            if (ex.InnerException is COMException innerComEx)
            {
                innerError += $" [COM: 0x{innerComEx.HResult:X8}]";
            }
        }

        return ex switch
        {
            PowerQueryCommandException pqEx => new ServiceResponse
            {
                Success = false,
                Command = command,
                SessionId = sessionId,
                ErrorCategory = pqEx.ErrorCategory,
                ErrorMessage = $"{pqEx.GetType().Name}: {SensitiveDataSanitizer.Redact(pqEx.Message)}",
                ExceptionType = exceptionType,
                HResult = hresult,
                InnerError = innerError
            },
            _ => new ServiceResponse
            {
                Success = false,
                Command = command,
                SessionId = sessionId,
                ErrorCategory = errorCategory,
                ErrorMessage = $"{exceptionType}: {SensitiveDataSanitizer.Redact(ex.Message)}",
                ExceptionType = exceptionType,
                HResult = hresult,
                InnerError = innerError
            }
        };
    }

    private ServiceResponse? TryBeginUsableSession(string sessionId, out IExcelBatch? batch)
    {
        if (!_sessionManager.TryBeginOperation(sessionId, out batch, out var errorMessage))
        {
            var (errorCategory, journalState) = ClassifySessionBeginFailure(errorMessage);
            if (journalState is not null && errorCategory is not null)
            {
                // Dead-process cleanup is handled synchronously by SessionManager's
                // DeadSessionCleanupStarting event before it removes tracking.
                if (!string.Equals(errorCategory, "ExcelProcessDied", StringComparison.Ordinal))
                {
                    _ = _safetyCoordinator.RecordSessionInterruption(sessionId, journalState, errorCategory);
                }
            }

            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = errorCategory,
                ErrorMessage = errorMessage
            };
        }

        return null;
    }

    private bool CloseInterruptedSession(string sessionId)
    {
        bool closed = false;
        try
        {
            closed = _sessionManager.CloseSession(sessionId, save: false, force: true) ||
                _sessionManager.GetSession(sessionId) is null;
        }
        catch (Exception cleanupEx)
        {
            System.Diagnostics.Debug.WriteLine($"Session cleanup failed for {sessionId}: {cleanupEx.Message}");
            try
            {
                closed = _sessionManager.GetSession(sessionId) is null;
            }
            catch (Exception verificationEx)
            {
                System.Diagnostics.Debug.WriteLine($"Session cleanup verification failed for {sessionId}: {verificationEx.Message}");
            }
        }

        if (closed)
        {
            _safetyCoordinator.RemoveSession(sessionId);
        }

        return closed;
    }

    private static (string? ErrorCategory, string? JournalState) ClassifySessionBeginFailure(string? errorMessage)
    {
        if (errorMessage?.Contains("has died", StringComparison.OrdinalIgnoreCase) == true ||
            errorMessage?.Contains("no longer running", StringComparison.OrdinalIgnoreCase) == true)
        {
            return ("ExcelProcessDied", "excelProcessDied");
        }

        if (errorMessage?.Contains("timed out or was cancelled", StringComparison.OrdinalIgnoreCase) == true)
        {
            return ("SessionInterrupted", "abortedUnknown");
        }

        return (null, null);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _shutdownCts.Cancel();
        try
        {
            foreach (var sessionId in _sessionManager.ActiveSessionIds.ToArray())
            {
                _safetyCoordinator.RecordServerShutdown(sessionId);
            }
        }
        finally
        {
            try
            {
                _sessionManager.Dispose();
            }
            finally
            {
                try
                {
                    _safetyCoordinator.Dispose();
                }
                finally
                {
                    _idempotencyCoordinator.Clear();
                    _shutdownCts.Dispose();
                }
            }
        }
    }
}

// === ARGUMENT TYPES (Session only - all other args are now generated in ServiceRegistry) ===

// Session
public sealed class SessionOpenArgs
{
    public string? FilePath { get; set; }
    public bool Show { get; set; }
    public int? TimeoutSeconds { get; set; }
}
public sealed class SessionCloseArgs { public bool Save { get; set; } }
public sealed class WorkflowOpenAndDescribeArgs
{
    public string? FilePath { get; set; }
    public bool Show { get; set; }
    public int? TimeoutSeconds { get; set; }
    public int? PreviewRows { get; set; }
    public int? PreviewColumns { get; set; }
}
public sealed class RecoveryArgs
{
    public string? RecoveryId { get; set; }
    public bool Show { get; set; }
    public int? TimeoutSeconds { get; set; }
}
