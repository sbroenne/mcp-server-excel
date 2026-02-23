using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.Diag;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Service.Rpc;
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
    private readonly SessionManager _sessionManager = new();
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

    public ExcelMcpService()
    {
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
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

                // Handle client via StreamJsonRpc — replaces hand-rolled JSON protocol
                // with standard JSON-RPC 2.0 over Content-Length-delimited framing.
                _ = Task.Run(async () =>
                {
                    await connectionLimit.WaitAsync(cancellationToken);
                    try
                    {
                        var rpcTarget = new DaemonRpcTarget(this);
                        using var rpc = JsonRpc.Attach(clientServer, rpcTarget);
                        await rpc.Completion; // Waits until client disconnects
                    }
                    finally
                    {
                        connectionLimit.Release();
                        try { if (clientServer.IsConnected) clientServer.Disconnect(); } catch { }
                        await clientServer.DisposeAsync();
                    }
                }, cancellationToken);
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
    public async Task<ServiceResponse> ProcessAsync(ServiceRequest request)
    {
        try
        {
            // Route command
            var parts = request.Command.Split('.', 2);
            var category = parts[0];
            var action = parts.Length > 1 ? parts[1] : "";

            return category switch
            {
                "service" => HandleServiceCommand(action),
                "session" => HandleSessionCommand(action, request),
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
                "diag" => DispatchSessionless(action, request),
                _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown command category: {category}" }
            };
        }
        catch (Exception ex)
        {
            // Include type name so callers can distinguish exception kinds (GitHub #482, Bug 5)
            return new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" };
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
        _shutdownCts.Cancel();
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

    private ServiceResponse HandleSessionCommand(string action, ServiceRequest request)
    {
        return action switch
        {
            "create" => HandleSessionCreate(request),
            "open" => HandleSessionOpen(request),
            "close" => HandleSessionClose(request),
            "save" => HandleSessionSave(request),
            "list" => HandleSessionList(),
            _ => new ServiceResponse { Success = false, ErrorMessage = $"Unknown session action: {action}" }
        };
    }

    private ServiceResponse HandleSessionCreate(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        var fullPath = Path.GetFullPath(args.FilePath);

        if (File.Exists(fullPath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"File already exists: {fullPath}. Use session open to open an existing workbook."
            };
        }

        var extension = Path.GetExtension(fullPath);
        if (!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Invalid file extension '{extension}'. session create supports .xlsx and .xlsm only."
            };
        }

        try
        {
            // Use the combined create+open which starts Excel only once
            TimeSpan? timeout = args.TimeoutSeconds.HasValue
                ? TimeSpan.FromSeconds(args.TimeoutSeconds.Value)
                : null;
            var sessionId = _sessionManager.CreateSessionForNewFile(fullPath, show: args.Show, operationTimeout: timeout, origin: SessionOrigin.CLI);

            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId, filePath = fullPath }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" };
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
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId, filePath = args.FilePath }, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" };
        }
    }

    private ServiceResponse HandleSessionClose(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var args = ServiceRegistry.DeserializeArgs<SessionCloseArgs>(request.Args);
        var closed = _sessionManager.CloseSession(request.SessionId, save: args?.Save ?? false);

        return closed
            ? new ServiceResponse { Success = true }
            : new ServiceResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
    }

    private ServiceResponse HandleSessionSave(ServiceRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var batch = _sessionManager.GetSession(request.SessionId);
        if (batch == null)
        {
            return new ServiceResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
        }

        // Check if Excel process is still alive before attempting save
        if (!batch.IsExcelProcessAlive())
        {
            _sessionManager.CloseSession(request.SessionId, save: false, force: true);
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Excel process for session '{request.SessionId}' has died. Session has been closed. Please create a new session."
            };
        }

        batch.Save();
        return new ServiceResponse { Success = true };
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



        return await WithSessionAsync(request.SessionId, batch => WrapResult(dispatch(action, batch)));

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

            // CopyToFile/MoveToFile are atomic operations without session

            if (sheetAction is SheetAction.CopyToFile or SheetAction.MoveToFile)

            {

                try

                {

                    return WrapResult(ServiceRegistry.Sheet.DispatchToCore(

                        _sheetCommands, sheetAction, null!, request.Args));

                }

                catch (Exception ex)

                {

                    return new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" };

                }

            }



            return await WithSessionAsync(request.SessionId, batch =>

                WrapResult(ServiceRegistry.Sheet.DispatchToCore(_sheetCommands, sheetAction, batch, request.Args)));

        }



        if (ServiceRegistry.SheetStyle.TryParseAction(actionString, out var styleAction))

        {

            return await WithSessionAsync(request.SessionId, batch =>

                WrapResult(ServiceRegistry.SheetStyle.DispatchToCore(_sheetCommands, styleAction, batch, request.Args)));

        }



        return new ServiceResponse { Success = false, ErrorMessage = $"Unknown sheet action: {actionString}" };

    }



    private async Task<ServiceResponse> DispatchRangeAsync(string actionString, ServiceRequest request)

    {

        return await WithSessionAsync(request.SessionId, batch =>

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

        return await WithSessionAsync(request.SessionId, batch =>

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

        return await WithSessionAsync(request.SessionId, batch =>
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


    private Task<ServiceResponse> WithSessionAsync(string? sessionId, Func<IExcelBatch, ServiceResponse> action)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return Task.FromResult(new ServiceResponse { Success = false, ErrorMessage = "sessionId is required" });
        }

        var batch = _sessionManager.GetSession(sessionId);
        if (batch == null)
        {
            return Task.FromResult(new ServiceResponse { Success = false, ErrorMessage = $"Session '{sessionId}' not found" });
        }

        // Check if Excel process is still alive before attempting operation
        if (!batch.IsExcelProcessAlive())
        {
            // Excel died - clean up the dead session
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Excel process for session '{sessionId}' has died. Session has been closed. Please create a new session."
            });
        }

        try
        {
            var response = action(batch);
            return Task.FromResult(response);
        }
        catch (TimeoutException ex)
        {
            // Operation timed out — Excel COM call is hung (IDispatch.Invoke stuck).
            // Force-close the session to trigger the force-kill path in ExcelBatch.Dispose(),
            // which will kill the hung Excel process and release the STA thread.
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Excel operation timed out and the session has been closed: {ex.Message} " +
                               "Please reopen the file with a new session."
            });
        }
        catch (OperationCanceledException)
        {
            // Caller cancelled (e.g., VS Code cancelled the tool call) while a COM operation
            // may still be running on the STA thread. ExcelBatch.Execute sets _operationTimedOut
            // on cancellation, but nobody calls Dispose() — the session stays alive with a
            // stuck STA thread, and all subsequent requests queue up and hang.
            // Force-close the session to kill the hung Excel process and release the STA thread.
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Operation was cancelled and the session has been closed. " +
                               "The Excel COM thread may have been unresponsive. " +
                               "Please reopen the file with a new session."
            });
        }
        catch (COMException ex) when (
            ex.HResult == ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE ||
            ex.HResult == ResiliencePipelines.RPC_E_CALL_FAILED)
        {
            // Excel process died during the operation — clean up the dead session
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Excel process for session '{sessionId}' has died (the application may have been closed or crashed). " +
                               "Session has been cleaned up. Please reopen the file with a new session."
            });
        }
        catch (InvalidOperationException ex) when (
            ex.Message.Contains("no longer running", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("process", StringComparison.OrdinalIgnoreCase))
        {
            // Excel process detected as dead before COM call (ExcelBatch pre-check)
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorMessage = $"Excel process for session '{sessionId}' is no longer running. " +
                               "Session has been cleaned up. Please reopen the file with a new session."
            });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new ServiceResponse { Success = false, ErrorMessage = $"{ex.GetType().Name}: {ex.Message}" });
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _shutdownCts.Cancel();
        _sessionManager.Dispose();
        _shutdownCts.Dispose();
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
