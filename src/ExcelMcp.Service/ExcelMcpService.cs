using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Service.Mac;
using Sbroenne.ExcelMcp.Core.Utilities;
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
    private readonly MacExcelBackend? _macBackend;
    private readonly MacExcelSessionManager? _macSessionManager;
    private readonly ConcurrentDictionary<string, byte> _knownSessionIds = new(StringComparer.Ordinal);
    private readonly DaemonHost _daemonHost;
    private readonly DateTime _startTime = DateTime.UtcNow;
    private bool _disposed;

    private readonly PlatformCommandSet _commands;
    private readonly FileCommands _fileCommands = new();

    public ExcelMcpService()
    {
        _commands = OperatingSystem.IsMacOS() ? PlatformCommandSet.CreateMac() : PlatformCommandSet.CreateWindows();
        _daemonHost = new DaemonHost(
            ProcessAsync,
            () => SessionCount);
        if (OperatingSystem.IsMacOS())
        {
            _macBackend = new MacExcelBackend();
            _macSessionManager = new MacExcelSessionManager(_macBackend);
        }
    }

    internal ExcelMcpService(MacExcelBackend macBackend)
    {
        _commands = PlatformCommandSet.CreateMac();
        _macBackend = macBackend;
        _macSessionManager = new MacExcelSessionManager(macBackend);
        _daemonHost = new DaemonHost(
            ProcessAsync,
            () => SessionCount);
    }

    public DateTime StartTime => _startTime;
    public int SessionCount => _macSessionManager?.Count ?? _sessionManager.GetActiveSessions().Count;
    public SessionManager SessionManager => _sessionManager;

    /// <summary>
    /// Runs the service in-process, listening for commands on the named pipe.
    /// This method blocks until shutdown is requested via <see cref="RequestShutdown"/>.
    /// </summary>
    /// <param name="pipeName">The named pipe to listen on.</param>
    /// <param name="idleTimeout">Optional idle timeout. Service shuts down after this duration with no active sessions. Null = no timeout.</param>
    public Task RunAsync(string pipeName, TimeSpan? idleTimeout = null) =>
        _daemonHost.RunAsync(pipeName, idleTimeout);

    public void RequestShutdown() => _daemonHost.RequestShutdown();

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

            ServiceRegistry.ValidateCommandArguments(request.Command, request.Args);

            ServiceResponse response = category switch
            {
                "service" => action is "helper-check" or "helper-build"
                    ? await HandleMacHelperCommandAsync(action, request)
                    : HandleServiceCommand(action),
                "session" => _macSessionManager is not null
                    ? await HandleMacSessionCommandAsync(action, request)
                    : HandleSessionCommand(action, request),
                "sheet" => await DispatchSheetAsync(action, request),
                "worksheetstyle" => await DispatchSimpleAsync<WorksheetStyleAction>(action, request,
                    ServiceRegistry.WorksheetStyle.TryParseAction,
                    (a, batch) => ServiceRegistry.WorksheetStyle.DispatchToCore(_commands.WorksheetStyle, a, batch, request.Args)),
                "range" or "rangeedit" or "rangeformat" or "rangelink" => await DispatchRangeAsync(action, request),
                "table" or "tablecolumn" => await DispatchTableAsync(action, request),
                "powerquery" => await DispatchSimpleAsync<PowerQueryAction>(action, request,
                    ServiceRegistry.PowerQuery.TryParseAction,
                    (a, batch) => ServiceRegistry.PowerQuery.DispatchToCore(_commands.PowerQuery, a, batch, request.Args)),
                "pivottable" => await DispatchSimpleAsync<PivotTableAction>(action, request,
                    ServiceRegistry.PivotTable.TryParseAction,
                    (a, batch) => ServiceRegistry.PivotTable.DispatchToCore(_commands.PivotTable, a, batch, request.Args)),
                "pivottablefield" => await DispatchSimpleAsync<PivotTableFieldAction>(action, request,
                    ServiceRegistry.PivotTableField.TryParseAction,
                    (a, batch) => ServiceRegistry.PivotTableField.DispatchToCore(_commands.PivotTableField, a, batch, request.Args)),
                "pivottablecalc" => await DispatchSimpleAsync<PivotTableCalcAction>(action, request,
                    ServiceRegistry.PivotTableCalc.TryParseAction,
                    (a, batch) => ServiceRegistry.PivotTableCalc.DispatchToCore(_commands.PivotTableCalc, a, batch, request.Args)),
                "chart" => await DispatchSimpleAsync<ChartAction>(action, request,
                    ServiceRegistry.Chart.TryParseAction,
                    (a, batch) => ServiceRegistry.Chart.DispatchToCore(_commands.Chart, a, batch, request.Args)),
                "chartconfig" => await DispatchSimpleAsync<ChartConfigAction>(action, request,
                    ServiceRegistry.ChartConfig.TryParseAction,
                    (a, batch) => ServiceRegistry.ChartConfig.DispatchToCore(_commands.ChartConfig, a, batch, request.Args)),
                "connection" => await DispatchSimpleAsync<ConnectionAction>(action, request,
                    ServiceRegistry.Connection.TryParseAction,
                    (a, batch) => ServiceRegistry.Connection.DispatchToCore(_commands.Connection, a, batch, request.Args)),
                "querytable" => await DispatchSimpleAsync<QueryTableAction>(action, request,
                    ServiceRegistry.QueryTable.TryParseAction,
                    (a, batch) => ServiceRegistry.QueryTable.DispatchToCore(_commands.QueryTable, a, batch, request.Args)),
                "calculationmode" => await DispatchSimpleAsync<CalculationModeAction>(action, request,
                    ServiceRegistry.CalculationMode.TryParseAction,
                    (a, batch) => ServiceRegistry.CalculationMode.DispatchToCore(_commands.CalculationMode, a, batch, request.Args)),
                "analysis" => await DispatchSimpleAsync<AnalysisAction>(action, request,
                    ServiceRegistry.Analysis.TryParseAction,
                    (a, batch) => ServiceRegistry.Analysis.DispatchToCore(_commands.Analysis, a, batch, request.Args)),
                "namedrange" => await DispatchSimpleAsync<NamedRangeAction>(action, request,
                    ServiceRegistry.NamedRange.TryParseAction,
                    (a, batch) => ServiceRegistry.NamedRange.DispatchToCore(_commands.NamedRange, a, batch, request.Args)),
                "conditionalformat" => await DispatchSimpleAsync<ConditionalFormatAction>(action, request,
                    ServiceRegistry.ConditionalFormat.TryParseAction,
                    (a, batch) => ServiceRegistry.ConditionalFormat.DispatchToCore(_commands.ConditionalFormat, a, batch, request.Args)),
                "vba" => await DispatchSimpleAsync<VbaAction>(action, request,
                    ServiceRegistry.Vba.TryParseAction,
                    (a, batch) => ServiceRegistry.Vba.DispatchToCore(_commands.Vba, a, batch, request.Args)),
                "datamodel" => await DispatchSimpleAsync<DataModelAction>(action, request,
                    ServiceRegistry.DataModel.TryParseAction,
                    (a, batch) => ServiceRegistry.DataModel.DispatchToCore(_commands.DataModel, a, batch, request.Args)),
                "datamodelrelationship" => await DispatchSimpleAsync<DataModelRelationshipAction>(action, request,
                    ServiceRegistry.DataModelRelationship.TryParseAction,
                    (a, batch) => ServiceRegistry.DataModelRelationship.DispatchToCore(_commands.DataModelRelationship, a, batch, request.Args)),
                "slicer" => await DispatchSimpleAsync<SlicerAction>(action, request,
                    ServiceRegistry.Slicer.TryParseAction,
                    (a, batch) => ServiceRegistry.Slicer.DispatchToCore(_commands.Slicer, a, batch, request.Args)),
                "screenshot" => await DispatchSimpleAsync<ScreenshotAction>(action, request,
                    ServiceRegistry.Screenshot.TryParseAction,
                    (a, batch) => ServiceRegistry.Screenshot.DispatchToCore(_commands.Screenshot, a, batch, request.Args)),
                "window" => await DispatchWindowAsync(action, request),
                "workbook" => await DispatchWorkbookAsync(action, request),
                "diag" => DispatchSessionless(action, request),
                "drawing" => await DispatchSimpleAsync<DrawingAction>(action, request,
                    ServiceRegistry.Drawing.TryParseAction,
                    (a, batch) => ServiceRegistry.Drawing.DispatchToCore(_commands.Drawing, a, batch, request.Args)),
                "pythoninexcel" => await DispatchSimpleAsync<PythonInExcelAction>(action, request,
                    ServiceRegistry.PythonInExcel.TryParseAction,
                    (a, batch) => ServiceRegistry.PythonInExcel.DispatchToCore(_commands.PythonInExcel, a, batch, request.Args)),
                "xmlmap" => await DispatchSimpleAsync<XmlMapAction>(action, request,
                    ServiceRegistry.XmlMap.TryParseAction,
                    (a, batch) => ServiceRegistry.XmlMap.DispatchToCore(_commands.XmlMap, a, batch, request.Args)),
                _ => new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "InvalidInput",
                    ErrorMessage = $"Unknown command category: {category}"
                }
            };

            return AttachRequestContext(request, response);
        }
        catch (MacExcelOperationException ex)
        {
            return AttachRequestContext(request, ex.ToServiceResponse());
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
            _ => new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown action '{action}' for command group 'service'. Valid actions: ping, shutdown, status."
            }
        };
    }

    private ServiceResponse HandleShutdown()
    {
        _daemonHost.RequestShutdownAfterResponse();
        return new ServiceResponse { Success = true };
    }

    private async Task<ServiceResponse> HandleMacHelperCommandAsync(string action, ServiceRequest request)
    {
        if (_macBackend == null)
        {
            throw new PlatformNotSupportedException("The native ExcelMcp helper requires Apple Silicon macOS.");
        }
        if (action == "helper-check")
        {
            if (ServiceRegistry.GetJsonPropertyNames(request.Args, includeNullValues: true).Count != 0)
            {
                throw new ArgumentException("service.helper-check accepts no arguments.");
            }
            var info = await _macBackend.RequireHelperAsync(MacHelperProtocol.InfoPrimitives, TimeSpan.FromSeconds(30));
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(info, ServiceProtocol.JsonOptions)
            };
        }
        var allowed = new HashSet<string>(["workbookPath", "outputPath", "helperVersion"], StringComparer.Ordinal);
        if (ServiceRegistry.GetJsonPropertyNames(request.Args, includeNullValues: true).Any(name => !allowed.Contains(name)))
        {
            throw new ArgumentException("service.helper-build accepts only workbookPath, outputPath, and helperVersion.");
        }
        var workbookPath = FilePathValidation.NormalizeAbsolutePath(GetRequiredStringArgument(request.Args, "workbookPath"));
        var outputPath = FilePathValidation.NormalizeAbsolutePath(GetRequiredStringArgument(request.Args, "outputPath"));
        var helperVersion = GetRequiredStringArgument(request.Args, "helperVersion");
        if (!string.Equals(Path.GetExtension(workbookPath), ".xlsm", StringComparison.OrdinalIgnoreCase)
            || !string.Equals(Path.GetExtension(outputPath), ".xlam", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("The helper build requires an Excel-authored .xlsm bootstrap and an .xlam output.");
        }
        if (!File.Exists(workbookPath)) throw new FileNotFoundException("The helper bootstrap workbook does not exist.", workbookPath);
        if (File.Exists(outputPath)) throw new InvalidOperationException("The helper build output already exists; it will not be overwritten.");
        var result = await _macBackend.InvokeAsync("helper.build", new { workbookPath, outputPath, helperVersion }, TimeSpan.FromMinutes(2));
        return new ServiceResponse { Success = true, Result = result.GetRawText() };
    }

    private ServiceResponse HandleStatus()
    {
        var status = new ServiceStatus
        {
            Running = true,
            ProcessId = Environment.ProcessId,
            SessionCount = SessionCount,
            StartTime = _startTime
        };
        return new ServiceResponse { Success = true, Result = JsonSerializer.Serialize(status, ServiceProtocol.JsonOptions) };
    }

    // === SESSION COMMANDS ===

    private async Task<ServiceResponse> HandleMacSessionCommandAsync(
        string action,
        ServiceRequest request)
    {
        if (action is not ("create" or "open" or "close" or "list" or "test"))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown session action: {action}"
            };
        }

        ValidateSessionActionArguments(action, request.Args);
        if (action == "list")
        {
            var sessions = _macSessionManager!.Sessions.Select(session => new
            {
                sessionId = session.SessionId,
                filePath = session.FilePath,
                isExcelVisible = session.IsVisible,
                activeOperations = Volatile.Read(ref session.ActiveOperations),
                canClose = session.PendingOperations == 0
                    && !session.IsClosing
                    && !session.HasUnconfirmedOpen
                    && !session.RequiresRecovery
                    && session.UnsafeReason is null,
                requiresRecovery = session.RequiresRecovery,
                unsafeReason = session.UnsafeReason
            }).ToList();
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(
                    new { success = true, sessions, count = sessions.Count },
                    ServiceProtocol.JsonOptions)
            };
        }

        if (action == "test")
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "PlatformNotSupported",
                ErrorMessage = MacCommandCapabilities.Get("file.test").UnavailableMessage
            };
        }

        if (action == "close")
        {
            if (string.IsNullOrWhiteSpace(request.SessionId))
            {
                return new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "InvalidInput",
                    ErrorMessage = "sessionId is required"
                };
            }

            var closeArgs = ServiceRegistry.DeserializeArgs<SessionCloseArgs>(request.Args);
            var closed = await _macSessionManager!.CloseAsync(request.SessionId, closeArgs?.Save ?? false);
            if (closed)
            {
                return new ServiceResponse { Success = true };
            }
            if (_knownSessionIds.ContainsKey(request.SessionId))
            {
                return new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(
                        new { success = true, sessionId = request.SessionId, message = "Session already closed." },
                        ServiceProtocol.JsonOptions)
                };
            }
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "SessionNotFound",
                ErrorMessage = $"Session '{request.SessionId}' not found"
            };
        }

        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args.FilePath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "filePath is required"
            };
        }

        var fullPath = FilePathValidation.NormalizeAbsolutePath(args.FilePath);
        var parsedTimeout = ParameterTransforms.ParseTimeoutSeconds(
            args.TimeoutSeconds,
            "timeoutSeconds",
            minimumSeconds: 10,
            maximumSeconds: 3600);
        var timeout = parsedTimeout ?? TimeSpan.FromSeconds(120);
        if (action == "open" && !File.Exists(fullPath))
        {
            throw new FileNotFoundException(
                $"Excel file not found: {fullPath}. To create a new file, use the 'create' action instead.",
                fullPath);
        }
        if (action == "create" && File.Exists(fullPath))
        {
            throw new InvalidOperationException(
                $"File already exists: {fullPath}. Use session open to open an existing workbook.");
        }

        var extension = Path.GetExtension(fullPath);
        if (!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"Invalid file extension '{extension}'. session {action} supports .xlsx and .xlsm only.");
        }

        var macroEnabled = string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase);
        var sessionId = action == "create"
            ? await _macSessionManager!.CreateAsync(fullPath, macroEnabled, args.Show, timeout)
            : await _macSessionManager!.OpenAsync(fullPath, args.Show, timeout);
        _knownSessionIds.TryAdd(sessionId, 0);
        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(
                new { success = true, sessionId, filePath = fullPath },
                ServiceProtocol.JsonOptions)
        };
    }

    private ServiceResponse HandleSessionCommand(string action, ServiceRequest request)
    {
        if (action is not ("create" or "open" or "close" or "list" or "test"))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown action '{action}' for command group 'session'. Valid actions: create, open, close, list, test."
            };
        }

        ValidateSessionActionArguments(action, request.Args);
        return action switch
        {
            "create" => HandleSessionCreate(request),
            "open" => HandleSessionOpen(request),
            "close" => HandleSessionClose(request),
            "list" => HandleSessionList(),
            "test" => HandleSessionTest(request),
            _ => throw new InvalidOperationException($"Unhandled session action: {action}")
        };
    }

    private static readonly Dictionary<string, string> SessionParameterDescriptions = new(StringComparer.Ordinal)
    {
        ["filePath"] = "absolute path to the workbook",
        ["show"] = "true to show the Excel window",
        ["timeoutSeconds"] = "whole seconds from 10 through 3600 (default 120)",
        ["save"] = "true to save the workbook before closing",
    };

    private static void ValidateSessionActionArguments(string action, string? argsJson)
    {
        string[] allowedParameters = action switch
        {
            "create" or "open" => ["filePath", "show", "timeoutSeconds"],
            "close" => ["save"],
            "test" => ["filePath", "timeoutSeconds"],
            _ => []
        };
        var unknownParameters = ServiceRegistry.GetJsonPropertyNames(argsJson, includeNullValues: true)
            .Where(parameter => !allowedParameters.Contains(parameter, StringComparer.Ordinal))
            .ToArray();
        if (unknownParameters.Length > 0)
        {
            var validParameters = allowedParameters.Length == 0
                ? "none"
                : string.Join("; ", allowedParameters.Select(parameter => $"{parameter} ({SessionParameterDescriptions[parameter]})"));
            throw new ArgumentException(
                $"Unknown parameter(s) for session.{action}: {string.Join(", ", unknownParameters)}. Valid parameters: {validParameters}.");
        }
    }

    private ServiceResponse HandleSessionCreate(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionOpenArgs>(request.Args);
        var timeout = ParameterTransforms.ParseTimeoutSeconds(
            args.TimeoutSeconds,
            "timeoutSeconds",
            minimumSeconds: 10,
            maximumSeconds: 3600);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "filePath is required"
            };
        }

        var fullPath = FilePathValidation.NormalizeAbsolutePath(args.FilePath);

        if (File.Exists(fullPath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Conflict",
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
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Invalid file extension '{extension}'. session create supports .xlsx and .xlsm only."
            };
        }

        try
        {
            // Use the combined create+open which starts Excel only once
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
        var timeout = ParameterTransforms.ParseTimeoutSeconds(
            args.TimeoutSeconds,
            "timeoutSeconds",
            minimumSeconds: 10,
            maximumSeconds: 3600);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "filePath is required"
            };
        }
        var fullPath = FilePathValidation.NormalizeAbsolutePath(args.FilePath);

        try
        {
            var sessionId = _sessionManager.CreateSession(fullPath, show: args.Show, operationTimeout: timeout, origin: SessionOrigin.CLI);
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

    private ServiceResponse HandleSessionClose(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionCloseArgs>(request.Args);
        var shouldSave = args?.Save ?? false;

        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "sessionId is required"
            };
        }

        bool closed;
        try
        {
            closed = _sessionManager.CloseSession(request.SessionId, save: shouldSave);
        }
        catch (Exception ex) when (IsFatalExcelDisconnect(ex))
        {
            CleanupDeadSession(request.SessionId);
            return CreateExcelDisconnectedResponse(request.SessionId, ex, shouldSave
                ? "Excel disconnected while saving before close. Session has been cleaned up; reopen the workbook and verify whether the save completed."
                : "Excel disconnected while closing. Session has been cleaned up; reopen the workbook with a new session.");
        }

        if (closed)
        {
            return new ServiceResponse { Success = true };
        }

        if (_knownSessionIds.ContainsKey(request.SessionId))
        {
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(
                    new { success = true, sessionId = request.SessionId, message = "Session already closed." },
                    ServiceProtocol.JsonOptions)
            };
        }

        return new ServiceResponse
        {
            Success = false,
            ErrorCategory = "SessionNotFound",
            ErrorMessage = $"Session '{request.SessionId}' not found"
        };
    }

    private ServiceResponse HandleSessionTest(ServiceRequest request)
    {
        var args = ServiceRegistry.DeserializeArgs<SessionTestArgs>(request.Args);
        var timeout = ParameterTransforms.ParseTimeoutSeconds(
            args.TimeoutSeconds,
            "timeoutSeconds",
            minimumSeconds: 10,
            maximumSeconds: 3600);
        if (string.IsNullOrWhiteSpace(args.FilePath))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "filePath is required"
            };
        }

        try
        {
            var result = _fileCommands.Test(args.FilePath);
            if (result.Exists
                && result.Extension is ".xlsx" or ".xlsm" or ".xlsb" or ".xls"
                && !result.IsIrmProtected
                && result.Message == null)
            {
                try
                {
                    _sessionManager.ValidateWorkbookOpen(result.FilePath, timeout);
                    result.IsValid = true;
                    result.CanOpen = true;
                }
                catch (Exception ex) when (ex is TimeoutException or OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    result.Message =
                        $"File is not a valid Excel workbook or Excel could not open it: {ex.Message}";
                }
            }

            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex);
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

            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown action: {actionString}"
            };



        return await WithSessionAsync(request, batch => WrapResult(dispatch(action, batch)));

    }

    /// <summary>
    /// Dispatches a session-less command (no Excel batch required).
    /// Used for [NoSession] categories like diag.
    /// </summary>
    private ServiceResponse DispatchSessionless(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.Diag.TryParseAction(actionString, out var action))
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown action: {actionString}"
            };

        return WrapResult(ServiceRegistry.Diag.DispatchToCore(_commands.Diag, action, request.Args));
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

                        _commands.Sheet, sheetAction, null!, request.Args));

                }

                catch (Exception ex)

                {

                    return CreateErrorResponse(ex);

                }

            }



            return await WithSessionAsync(request, batch =>

                WrapResult(ServiceRegistry.Sheet.DispatchToCore(_commands.Sheet, sheetAction, batch, request.Args)));

        }



        return new ServiceResponse
        {
            Success = false,
            ErrorCategory = "InvalidInput",
            ErrorMessage = $"Unknown sheet action: {actionString}"
        };

    }



    private async Task<ServiceResponse> DispatchRangeAsync(string actionString, ServiceRequest request)
    {
        return await WithSessionAsync(request, batch =>
        {
            if (ServiceRegistry.Range.TryParseAction(actionString, out var ra))
                return WrapResult(ServiceRegistry.Range.DispatchToCore(_commands.Range, ra, batch, request.Args));

            if (ServiceRegistry.RangeEdit.TryParseAction(actionString, out var rea))
                return WrapResult(ServiceRegistry.RangeEdit.DispatchToCore(_commands.RangeEdit, rea, batch, request.Args));

            if (ServiceRegistry.RangeFormat.TryParseAction(actionString, out var rfa))
                return WrapResult(ServiceRegistry.RangeFormat.DispatchToCore(_commands.RangeFormat, rfa, batch, request.Args));

            if (ServiceRegistry.RangeLink.TryParseAction(actionString, out var rla))
                return WrapResult(ServiceRegistry.RangeLink.DispatchToCore(_commands.RangeLink, rla, batch, request.Args));

            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown range action: {actionString}"
            };
        });
    }


    private async Task<ServiceResponse> DispatchTableAsync(string actionString, ServiceRequest request)

    {

        return await WithSessionAsync(request, batch =>

        {

            if (ServiceRegistry.Table.TryParseAction(actionString, out var ta))

                return WrapResult(ServiceRegistry.Table.DispatchToCore(_commands.Table, ta, batch, request.Args));

            if (ServiceRegistry.TableColumn.TryParseAction(actionString, out var tca))

                return WrapResult(ServiceRegistry.TableColumn.DispatchToCore(_commands.TableColumn, tca, batch, request.Args));

            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown table action: {actionString}"
            };

        });

    }

    private async Task<ServiceResponse> DispatchWindowAsync(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.Window.TryParseAction(actionString, out var windowAction))
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown window action: {actionString}"
            };

        return await WithSessionAsync(request, batch =>
        {
            var result = WrapResult(ServiceRegistry.Window.DispatchToCore(_commands.Window, windowAction, batch, request.Args));

            // Update SessionManager visibility flag when show/hide commands succeed
            if (_macSessionManager == null && result.Success && !string.IsNullOrWhiteSpace(request.SessionId))
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

    private async Task<ServiceResponse> DispatchWorkbookAsync(string actionString, ServiceRequest request)
    {
        if (!ServiceRegistry.Workbook.TryParseAction(actionString, out var workbookAction))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown workbook action: {actionString}"
            };
        }

        return await WithSessionAsync(request, batch =>
        {
            string? reservedPath = null;
            var releaseReservation = true;
            if (_macSessionManager == null && workbookAction == WorkbookAction.SaveAs &&
                !string.IsNullOrWhiteSpace(request.SessionId))
            {
                reservedPath = _sessionManager.ReserveSessionFilePath(
                    request.SessionId,
                    GetRequiredStringArgument(request.Args, "targetPath"));
            }

            try
            {
                var result = WrapResult(
                    ServiceRegistry.Workbook.DispatchToCore(_commands.Workbook, workbookAction, batch, request.Args));

                return result;
            }
            catch (Exception ex) when (ex is TimeoutException or OperationCanceledException)
            {
                // The COM call may still be mutating the target. WithSessionAsync force-closes
                // the session, whose cleanup releases every path claim after Excel terminates.
                releaseReservation = false;
                throw;
            }
            finally
            {
                if (reservedPath != null && releaseReservation)
                {
                    if (string.Equals(batch.WorkbookPath, reservedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        _sessionManager.UpdateSessionFilePath(request.SessionId!, batch.WorkbookPath);
                    }
                    _sessionManager.ReleaseSessionFilePathReservation(request.SessionId!, reservedPath);
                }
            }
        });
    }

    private static string GetRequiredStringArgument(string? args, string argumentName)
    {
        if (string.IsNullOrWhiteSpace(args))
        {
            throw new ArgumentException($"{argumentName} is required.", argumentName);
        }

        using var document = JsonDocument.Parse(args);
        foreach (var property in document.RootElement.EnumerateObject())
        {
            if (string.Equals(property.Name, argumentName, StringComparison.OrdinalIgnoreCase) &&
                property.Value.ValueKind == JsonValueKind.String &&
                !string.IsNullOrWhiteSpace(property.Value.GetString()))
            {
                return property.Value.GetString()!;
            }
        }

        throw new ArgumentException($"{argumentName} is required.", argumentName);
    }


    private Task<ServiceResponse> WithSessionAsync(ServiceRequest request, Func<IExcelBatch, ServiceResponse> action)
    {
        var sessionId = request.SessionId;
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "sessionId is required"
            });
        }

        if (_macSessionManager != null)
        {
            return WithMacSessionAsync(sessionId, action);
        }

        var sessionError = TryBeginUsableSession(sessionId, out var batch);
        if (sessionError != null)
        {
            return Task.FromResult(sessionError);
        }

        try
        {
            ServiceRegistry.ValidateWorkbookWriteAccess(request.Command, batch!, request.Args);
            var response = action(batch!);
            return Task.FromResult(response);
        }
        catch (TimeoutException ex)
        {
            if (batch!.HasTimedOutOperation)
            {
                CleanupDeadSession(sessionId);
                return Task.FromResult(new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "Timeout",
                    ErrorMessage = $"Excel operation timed out after execution started and the session has been closed: {ex.Message} " +
                                   "Please reopen the file with a new session.",
                    ExceptionType = ex.GetType().Name
                });
            }

            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Timeout",
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name
            });
        }
        catch (OperationCanceledException ex)
        {
            if (batch!.HasTimedOutOperation)
            {
                CleanupDeadSession(sessionId);
                return Task.FromResult(new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "Cancelled",
                    ErrorMessage = "Operation was cancelled after execution started and the session has been closed. " +
                                   "The Excel COM thread may have been unresponsive. " +
                                   "Please reopen the file with a new session.",
                    ExceptionType = ex.GetType().Name
                });
            }

            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Cancelled",
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name
            });
        }
        catch (COMException ex) when (
            ex.HResult == ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE ||
            ex.HResult == ResiliencePipelines.RPC_E_CALL_FAILED ||
            ex.HResult == ResiliencePipelines.RPC_E_DISCONNECTED)
        {
            // Excel process died during the operation — clean up the dead session
            CleanupDeadSession(sessionId);
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
        catch (Exception ex)
        {
            if (IsFatalExcelDisconnect(ex))
            {
                CleanupDeadSession(sessionId);
                return Task.FromResult(CreateExcelDisconnectedResponse(sessionId, ex,
                    $"Excel process for session '{sessionId}' disconnected during the operation. Session has been cleaned up. Please reopen the file with a new session."));
            }

            // Check if Excel died with a non-COM exception — clean up dead session
            if (batch != null && !batch.IsExcelProcessAlive())
            {
                CleanupDeadSession(sessionId);
                return Task.FromResult(CreateExcelDisconnectedResponse(sessionId, ex,
                    $"Excel process for session '{sessionId}' is no longer running. Session has been cleaned up. Please reopen the file with a new session."));
            }

            return Task.FromResult(CreateErrorResponse(ex));
        }
        finally
        {
            _sessionManager.EndOperation(sessionId);
        }
    }

    private async Task<ServiceResponse> WithMacSessionAsync(string sessionId, Func<IExcelBatch, ServiceResponse> action)
    {
        try
        {
            return await _macSessionManager!.ExecuteAsync(sessionId, async session =>
            {
                using var batch = new MacExcelBatch(_macBackend!, session);
                return await Task.Run(() => action(batch));
            });
        }
        catch (KeyNotFoundException ex)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "SessionNotFound",
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name
            };
        }
        catch (TimeoutException ex)
        {
            _macSessionManager!.RequireRecovery(sessionId);
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Timeout",
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name
            };
        }
        catch (MacExcelOperationException ex)
        {
            return ex.ToServiceResponse();
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex);
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

        if (string.Equals(command, response.Command, StringComparison.Ordinal)
            && string.Equals(sessionId, response.SessionId, StringComparison.Ordinal))
        {
            return response;
        }

        return CloneResponse(response, command, sessionId);
    }

    private static ServiceResponse CloneResponse(ServiceResponse response, string? command, string? sessionId)
    {
        return new ServiceResponse
        {
            Success = response.Success,
            Command = command,
            SessionId = sessionId,
            ErrorMessage = response.ErrorMessage,
            ErrorCategory = response.ErrorCategory,
            ExceptionType = response.ExceptionType,
            HResult = response.HResult,
            InnerError = response.InnerError,
            Result = response.Result
        };
    }

    private static ServiceResponse CreateErrorResponse(Exception ex, string? command = null, string? sessionId = null)
    {
        var exceptionType = ex.GetType().Name;
        string? hresult = OperationFailureClassifier.GetComHResult(ex);
        string? innerError = null;
        var errorCategory = OperationFailureClassifier.Classify(ex);

        if (ex.InnerException != null)
        {
            innerError = ex.InnerException.Message;
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
                ErrorMessage = $"{pqEx.GetType().Name}: {pqEx.Message}",
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
                ErrorMessage = $"{exceptionType}: {ex.Message}",
                ExceptionType = exceptionType,
                HResult = hresult,
                InnerError = innerError
            }
        };
    }

    private ServiceResponse? TryBeginUsableSession(string sessionId, out IExcelBatch? batch)
    {
        if (!_sessionManager.TryBeginOperation(
            sessionId,
            out batch,
            out var errorMessage,
            out var sessionError))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = sessionError switch
                {
                    SessionOperationError.MissingSessionId => "InvalidInput",
                    SessionOperationError.NotFound => "SessionNotFound",
                    SessionOperationError.Closing
                        or SessionOperationError.Quarantined => "SessionUnavailable",
                    SessionOperationError.TimedOutOrCancelled => "SessionInvalidated",
                    SessionOperationError.ExcelProcessDied => "ExcelProcessDied",
                    _ => null
                },
                ErrorMessage = errorMessage
            };
        }

        return null;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _daemonHost.RequestShutdown();
        try
        {
            _macSessionManager?.Dispose();
            _sessionManager.Dispose();
        }
        finally
        {
            _daemonHost.Dispose();
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
public sealed class SessionTestArgs
{
    public string? FilePath { get; set; }
    public int? TimeoutSeconds { get; set; }
}
