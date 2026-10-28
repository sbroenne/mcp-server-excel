using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Analysis;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.Diag;
using Sbroenne.ExcelMcp.Core.Commands.Drawing;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.PythonInExcel;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;
using Sbroenne.ExcelMcp.Core.Commands.Settings;
using Sbroenne.ExcelMcp.Core.Commands.Window;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Sbroenne.ExcelMcp.Core.Commands.XmlMap;
using Sbroenne.ExcelMcp.Core.Models;
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
    private readonly ConcurrentDictionary<string, byte> _knownSessionIds = new(StringComparer.Ordinal);
    private readonly IPowerQueryTraceSettings _pqTraceSettings;
    private readonly DaemonHost _daemonHost;
    private readonly Lock _shutdownLock = new();
    private readonly DateTime _startTime = DateTime.UtcNow;
    private bool _disposed;
    private volatile bool _shuttingDown;
    private string? _shutdownReport;

    // Core command instances - use concrete types per CA1859
    private readonly RangeCommands _rangeCommands = new();
    private readonly SheetCommands _sheetCommands = new();
    private readonly TableCommands _tableCommands = new();
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly PivotTableCommands _pivotTableCommands = new();
    private readonly SlicerCommands _slicerCommands = new();
    private readonly ChartCommands _chartCommands = new();
    private readonly ConnectionCommands _connectionCommands = new();
    private readonly QueryTableCommands _queryTableCommands = new();
    private readonly NamedRangeCommands _namedRangeCommands = new();
    private readonly ConditionalFormattingCommands _conditionalFormatCommands = new();
    private readonly VbaCommands _vbaCommands = new();
    private readonly DataModelCommands _dataModelCommands = new();
    private readonly CalculationModeCommands _calculationModeCommands = new();
    private readonly ScreenshotCommands _screenshotCommands = new();
    private readonly DiagCommands _diagCommands = new();
    private readonly SettingsCommands _settingsCommands;
    private readonly PowerQueryTraceCommands _powerQueryTraceCommands;
    private readonly DrawingCommands _drawingCommands = new();
    private readonly WindowCommands _windowCommands = new();
    private readonly WorkbookCommands _workbookCommands = new();
    private readonly PythonInExcelCommands _pythonInExcelCommands = new();
    private readonly AnalysisCommands _analysisCommands = new();
    private readonly XmlMapCommands _xmlMapCommands = new();
    private readonly FileCommands _fileCommands = new();

    public ExcelMcpService() : this(null)
    {
    }

    internal ExcelMcpService(
        Func<ServiceRequest, Func<Task<ServiceResponse>>, Task<ServiceResponse>>? trackDaemonRequest,
        IPowerQueryTraceSettings? pqTraceSettings = null)
    {
        _pqTraceSettings = pqTraceSettings ?? new RegistryPowerQueryTraceSettings();
        _settingsCommands = new SettingsCommands(_pqTraceSettings);
        _powerQueryTraceCommands = new PowerQueryTraceCommands(
            _pqTraceSettings.TraceFolder,
            ResolveSessionProcess);
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands, _pqTraceSettings);
        _daemonHost = new DaemonHost(
            request => trackDaemonRequest == null
                ? ProcessAsync(request)
                : trackDaemonRequest(request, () => ProcessAsync(request)),
            requestAbandoned: RevealAbandonedRequest);
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A caller-abandonment notification only reveals retained ownership. Failures are diagnosed without replacing cancellation or closing a workbook whose execution is still owned.")]
    internal void RevealAbandonedRequest(ServiceRequest request, ServiceResponse? response)
    {
        try
        {
            var sessionId = response?.SessionId ?? request.SessionId;
            if (string.IsNullOrWhiteSpace(sessionId)
                && request.Command is "session.open" or "session.create"
                && response is { Success: true })
            {
                using var result = JsonDocument.Parse(response.Result
                    ?? throw new InvalidOperationException("Session creation returned no result."));
                sessionId = result.RootElement.GetProperty("sessionId").GetString();
                if (string.IsNullOrWhiteSpace(sessionId))
                    throw new InvalidOperationException("Session creation returned no session ID.");
            }
            if (!string.IsNullOrWhiteSpace(sessionId))
                _sessionManager.RevealForRecovery(sessionId, waitExpired: true);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.TraceError(
                "Could not reveal retained Excel ownership after the caller abandoned {0}: {1}",
                request.Command, ex);
        }
    }

    public DateTime StartTime => _startTime;
    public int SessionCount => _sessionManager.GetActiveSessions(statusOnly: true).Count;

    internal static bool IsSessionlessCommand(string command) =>
        command.Split('.', 2)[0] is "service" or "diag" or "settings" or "powerquerytrace"
        || command is "session.open" or "session.create" or "session.test"
            or "sheet.copy-to-file" or "sheet.move-to-file";
    public SessionManager SessionManager => _sessionManager;

    /// <summary>
    /// Runs the service in-process, listening for commands on the named pipe.
    /// This method blocks until shutdown is requested via <see cref="RequestShutdown"/>.
    /// </summary>
    /// <param name="pipeName">The named pipe to listen on.</param>
    public Task RunAsync(string pipeName) =>
        _daemonHost.RunAsync(pipeName);

    /// <summary>
    /// Shuts down Excel sessions and stops the service, even when an Excel process cannot be
    /// confirmed. Returns null, or a report of each Excel process left as it is.
    /// </summary>
    public string? RequestShutdown()
    {
        lock (_shutdownLock)
        {
            var report = ShutDownSessions();
            _daemonHost.RequestShutdown();
            return report;
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Shutdown is a failure-transport boundary: each session failure, with its Excel process, is returned to the caller as a report while the service still stops. Excel is left as it is.")]
    private string? ShutDownSessions()
    {
        if (_shuttingDown)
            return _shutdownReport;
        _shuttingDown = true;
        try
        {
            _sessionManager.Dispose();
            return null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.TraceWarning("Excel shutdown was not confirmed: {0}", ex);
            IEnumerable<Exception> failures = ex is AggregateException aggregate ? aggregate.InnerExceptions : [ex];
            return _shutdownReport = "The service stopped, but some Excel processes were left as they are; nothing was saved, " +
                "discarded, or terminated. " +
                string.Join(" ", failures.Select(failure => failure.Message));
        }
    }

    /// <summary>
    /// Processes a service request directly (in-process, no pipe).
    /// Product controllers use this engine inside their private Excel workers.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Service request failure-transport boundary returns CreateErrorResponse with Success=false, original exception category/message/HResult, and request/session context.")]
    public async Task<ServiceResponse> ProcessAsync(ServiceRequest request)
    {
        try
        {
            request.CallerCancellation.ThrowIfCancellationRequested();
            // Route command
            var parts = request.Command.Split('.', 2);
            var category = parts[0];
            var action = parts.Length > 1 ? parts[1] : "";

            ServiceRegistry.ValidateCommandArguments(request.Command, request.Args);

            if (_shuttingDown && category != "service")
            {
                return AttachRequestContext(request, new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "SessionUnavailable",
                    ErrorMessage = "The service is shutting down and accepts no new work. Nothing was changed."
                });
            }

            ServiceResponse response = category switch
            {
                "service" => HandleServiceCommand(action),
                "session" => await HandleSessionCommandAsync(action, request),
                "sheet" => await DispatchSheetAsync(action, request),
                "worksheetstyle" => await DispatchSimpleAsync<WorksheetStyleAction>(action, request,
                    ServiceRegistry.WorksheetStyle.TryParseAction,
                    (a, batch) => ServiceRegistry.WorksheetStyle.DispatchToCore(_sheetCommands, a, batch, request.Args)),
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
                "querytable" => await DispatchSimpleAsync<QueryTableAction>(action, request,
                    ServiceRegistry.QueryTable.TryParseAction,
                    (a, batch) => ServiceRegistry.QueryTable.DispatchToCore(_queryTableCommands, a, batch, request.Args)),
                "calculationmode" => await DispatchSimpleAsync<CalculationModeAction>(action, request,
                    ServiceRegistry.CalculationMode.TryParseAction,
                    (a, batch) => ServiceRegistry.CalculationMode.DispatchToCore(_calculationModeCommands, a, batch, request.Args)),
                "analysis" => await DispatchSimpleAsync<AnalysisAction>(action, request,
                    ServiceRegistry.Analysis.TryParseAction,
                    (a, batch) => ServiceRegistry.Analysis.DispatchToCore(_analysisCommands, a, batch, request.Args)),
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
                "datamodelrelationship" => await DispatchSimpleAsync<DataModelRelationshipAction>(action, request,
                    ServiceRegistry.DataModelRelationship.TryParseAction,
                    (a, batch) => ServiceRegistry.DataModelRelationship.DispatchToCore(_dataModelCommands, a, batch, request.Args)),
                "slicer" => await DispatchSimpleAsync<SlicerAction>(action, request,
                    ServiceRegistry.Slicer.TryParseAction,
                    (a, batch) => ServiceRegistry.Slicer.DispatchToCore(_slicerCommands, a, batch, request.Args)),
                "screenshot" => await DispatchSimpleAsync<ScreenshotAction>(action, request,
                    ServiceRegistry.Screenshot.TryParseAction,
                    (a, batch) => ServiceRegistry.Screenshot.DispatchToCore(_screenshotCommands, a, batch, request.Args)),
                "window" => await DispatchWindowAsync(action, request),
                "workbook" => await DispatchWorkbookAsync(action, request),
                "diag" or "settings" or "powerquerytrace" => DispatchSessionless(category, action, request),
                "drawing" => await DispatchSimpleAsync<DrawingAction>(action, request,
                    ServiceRegistry.Drawing.TryParseAction,
                    (a, batch) => ServiceRegistry.Drawing.DispatchToCore(_drawingCommands, a, batch, request.Args)),
                "pythoninexcel" => await DispatchSimpleAsync<PythonInExcelAction>(action, request,
                    ServiceRegistry.PythonInExcel.TryParseAction,
                    (a, batch) => ServiceRegistry.PythonInExcel.DispatchToCore(_pythonInExcelCommands, a, batch, request.Args)),
                "xmlmap" => await DispatchSimpleAsync<XmlMapAction>(action, request,
                    ServiceRegistry.XmlMap.TryParseAction,
                    (a, batch) => ServiceRegistry.XmlMap.DispatchToCore(_xmlMapCommands, a, batch, request.Args)),
                _ => new ServiceResponse
                {
                    Success = false,
                    ErrorCategory = "InvalidInput",
                    ErrorMessage = $"Unknown command category: {category}"
                }
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
        lock (_shutdownLock)
        {
            var report = ShutDownSessions();
            _daemonHost.RequestShutdownAfterResponse();
            return new ServiceResponse { Success = true, PartialState = report };
        }
    }

    private ServiceResponse HandleStatus()
    {
        var status = new ServiceStatus
        {
            Running = true,
            ProcessId = Environment.ProcessId,
            SessionCount = _sessionManager.GetActiveSessions(statusOnly: true).Count,
            StartTime = _startTime
        };
        return new ServiceResponse { Success = true, Result = JsonSerializer.Serialize(status, ServiceProtocol.JsonOptions) };
    }

    // === SESSION COMMANDS ===

    private async Task<ServiceResponse> HandleSessionCommandAsync(string action, ServiceRequest request)
    {
        if (action is not ("create" or "open" or "close" or "save" or "list" or "test" or "discover" or "attach"
            or "notices" or "enable-editing" or "enable-connections" or "enable-macros"))
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown action '{action}' for command group 'session'. Valid actions: create, open, close, save, list, test, discover, attach, notices, enable-editing, enable-connections, enable-macros."
            };
        }

        ValidateSessionActionArguments(action, request.Args);
        return action switch
        {
            "create" => HandleSessionCreate(request),
            "open" => HandleSessionOpen(request),
            "close" => HandleSessionClose(request),
            "save" => HandleSessionSave(request),
            "list" => HandleSessionList(),
            "test" => HandleSessionTest(request),
            "discover" => await HandleSessionDiscoverAsync(),
            "attach" => HandleSessionAttach(request),
            "notices" or "enable-editing" or "enable-connections" or "enable-macros" =>
                await HandleSessionSecurityAsync(action, request),
            _ => throw new InvalidOperationException($"Unhandled session action: {action}")
        };
    }

    private Task<ServiceResponse> HandleSessionSecurityAsync(string action, ServiceRequest request) =>
        WithSessionAsync(request, batch =>
        {
            var owner = batch is AttachedWorkbookBatch attached ? attached.Owner
                : batch as ExcelBatch ?? throw new InvalidOperationException("Workbook inspection is unavailable for this session.");
            var workbookId = (batch as AttachedWorkbookBatch)?.WorkbookId;
            if (action == "notices")
                return WrapResult(JsonSerializer.Serialize(
                    new { success = true, notices = owner.ReadWorkbookNotices(workbookId) }, ServiceProtocol.JsonOptions));
            using var args = JsonDocument.Parse(request.Args ?? "{}");
            bool ReadApproval(string name) =>
                args.RootElement.TryGetProperty(name, out var value) && value.GetBoolean();
            var result = owner.ApplyWorkbookSecurity(action, ReadApproval("approved"),
                ReadApproval("allowRememberedTrust"), ReadApproval("allowReopen"), workbookId);
            return WrapResult(JsonSerializer.Serialize(new { success = true, security = result }, ServiceProtocol.JsonOptions));
        });

    private async Task<ServiceResponse> HandleSessionDiscoverAsync()
    {
        var discovery = await _sessionManager.DiscoverOpenWorkbooksAsync();
        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                discovery.Workbooks,
                discovery.Errors,
                discovery.Complete
            }, ServiceProtocol.JsonOptions)
        };
    }

    private ServiceResponse HandleSessionAttach(ServiceRequest request)
    {
        using var args = JsonDocument.Parse(request.Args ?? "{}");
        if (!args.RootElement.TryGetProperty("workbookId", out var selection)
            || selection.ValueKind != JsonValueKind.String
            || string.IsNullOrWhiteSpace(selection.GetString()))
        {
            throw new ArgumentException("workbookId from session discover is required.");
        }

        var attached = _sessionManager.AttachWorkbook(selection.GetString()!);
        _knownSessionIds.TryAdd(attached.SessionId, 0);
        var warning = ExcelCapabilities.AutoSaveWarning(attached.AutoSaveOn);
        return new ServiceResponse
        {
            Success = true,
            SessionId = attached.SessionId,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                attached.SessionId,
                attached.FilePath,
                attached.Name,
                attached.AutoSaveOn,
                warning
            }, ServiceProtocol.JsonOptions)
        };
    }

    private ServiceResponse HandleSessionSave(ServiceRequest request) =>
        WithSessionAsync(request, batch =>
        {
            batch.Save();
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, sessionId = request.SessionId, saved = true },
                    ServiceProtocol.JsonOptions)
            };
        }).GetAwaiter().GetResult();

    private static readonly Dictionary<string, string> SessionParameterDescriptions = new(StringComparer.Ordinal)
    {
        ["filePath"] = "absolute path to the workbook",
        ["show"] = "true to show the Excel window",
        ["timeoutSeconds"] = "whole seconds from 10 through 3600 (default 120)",
        ["save"] = "true to save the workbook before closing",
        ["workbookId"] = "an open-workbook identifier returned by session discover",
        ["approved"] = "true only after explicit approval of this action on the selected workbook",
        ["allowRememberedTrust"] = "true only after approval that Excel can remember this file as trusted",
        ["allowReopen"] = "true only after separate approval to close and reopen the saved workbook; opening macros can run",
    };

    internal static void ValidateSessionActionArguments(string action, string? argsJson)
    {
        string[] allowedParameters = action switch
        {
            "create" or "open" => ["filePath", "show", "timeoutSeconds"],
            "close" => ["save"],
            "test" => ["filePath", "timeoutSeconds"],
            "attach" => ["workbookId"],
            "enable-editing" or "enable-connections" => ["approved", "allowRememberedTrust"],
            "enable-macros" => ["approved", "allowReopen"],
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
        if (action is "enable-editing" or "enable-connections" or "enable-macros")
        {
            using var args = JsonDocument.Parse(argsJson ?? "{}");
            foreach (var parameter in args.RootElement.EnumerateObject())
                if (parameter.Value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
                    throw new ArgumentException($"{parameter.Name} must be true or false for session.{action}.");
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Session creation transports failures through CreateErrorResponse, never a successful session response.")]
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

        var fullPath = FilePathValidation.NormalizeAbsoluteWindowsPath(args.FilePath);

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
            return CreateSessionOpenResponse(sessionId, fullPath, request.Command);
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Session open transports the original failure through CreateErrorResponse with Success=false.")]
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
        var fullPath = FilePathValidation.NormalizeWorkbookLocation(args.FilePath);

        try
        {
            var sessionId = _sessionManager.CreateSession(args.FilePath, show: args.Show, operationTimeout: timeout, origin: SessionOrigin.CLI);
            _knownSessionIds.TryAdd(sessionId, 0);
            return CreateSessionOpenResponse(sessionId, fullPath, request.Command);
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex);
        }
    }

    private ServiceResponse CreateSessionOpenResponse(string sessionId, string filePath, string command) =>
        WithSessionAsync(new ServiceRequest { Command = command, SessionId = sessionId }, batch =>
        {
            var autoSaveOn = batch.Execute((context, _) =>
                ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn));
            return new ServiceResponse
            {
                Success = true,
                SessionId = sessionId,
                Result = JsonSerializer.Serialize(new
                {
                    success = true,
                    sessionId,
                    filePath,
                    autoSaveOn,
                    warning = ExcelCapabilities.AutoSaveWarning(autoSaveOn)
                }, ServiceProtocol.JsonOptions)
            };
        }).GetAwaiter().GetResult();

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Only confirmed fatal Excel disconnections reach this handler; CreateExcelDisconnectedResponse reports failure and the original COM error after dead-session cleanup.")]
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
            _sessionManager.SetOperationContext(request.SessionId, request.Command, request.CallerCancellation);
            closed = _sessionManager.CloseSession(request.SessionId, save: shouldSave);
        }
        catch (Exception ex) when (IsFatalExcelDisconnect(ex))
        {
            return HandleExcelDisconnect(request.SessionId, ex);
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

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "This diagnostic action reports open failures in the validation result Message with IsValid/CanOpen false; other failures use CreateErrorResponse, and cancellation/timeouts propagate to it.")]
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
                && !result.IsIrmProtected
                && result.Message == null)
            {
                try
                {
                    _sessionManager.ValidateWorkbookOpen(result.FilePath, timeout);
                    result.IsValid = true;
                    result.CanOpen = true;
                }
                catch (Exception ex) when (ex is TimeoutException or OperationCanceledException or SessionRetainedException)
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
            return CreateErrorResponse(ex, request.Command, request.SessionId);
        }
    }

    private ServiceResponse HandleSessionList()
    {
        var sessions = _sessionManager.GetActiveSessions(statusOnly: true)
            .Select(s =>
            {
                var validation = _sessionManager.ValidateClose(s.SessionId);
                var operation = _sessionManager.GetOperationStatus(s.SessionId, reconcilePaths: false);
                return new SessionListEntry
                {
                    SessionId = s.SessionId,
                    FilePath = s.FilePath,
                    FilePaths = _sessionManager.GetOwnedFilePaths(s.SessionId),
                    WorkbookName = s.WorkbookName,
                    IsAttached = s.IsAttached,
                    IsExcelVisible = validation.IsExcelVisible,
                    ActiveOperations = validation.ActiveOperationCount,
                    OperationPending = validation.OperationPending || operation?.ExecutionPending == true,
                    CanClose = validation.CanClose,
                    ExcelState = validation.ExcelState,
                    BlockingReason = validation.BlockingReason,
                    BlockingMessage = validation.BlockingMessage,
                    Operation = operation
                };
            })
            .ToList();

        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new SessionListResult { Success = true, Sessions = sessions }, ServiceProtocol.JsonOptions)
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Sessionless command failures are transported through CreateErrorResponse with Success=false and the original exception, command and session context.")]
    private ServiceResponse DispatchSessionless(string category, string actionString, ServiceRequest request)
    {
        try
        {
            if (category == "diag")
            {
                if (ServiceRegistry.Diag.TryParseAction(actionString, out var action))
                {
                    return WrapResult(ServiceRegistry.Diag.DispatchToCore(_diagCommands, action, request.Args));
                }
            }
            else if (category == "settings")
            {
                if (ServiceRegistry.Settings.TryParseAction(actionString, out var action))
                {
                    return WrapResult(ServiceRegistry.Settings.DispatchToCore(_settingsCommands, action, request.Args));
                }
            }
            else if (category == "powerquerytrace")
            {
                if (ServiceRegistry.PowerQueryTrace.TryParseAction(actionString, out var action))
                {
                    return WrapResult(ServiceRegistry.PowerQueryTrace.DispatchToCore(_powerQueryTraceCommands, action, request.Args));
                }
            }

            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = $"Unknown action: {actionString}"
            };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex, request.Command, request.SessionId);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Atomic sessionless sheet operations transport their exact failure through CreateErrorResponse with Success=false.")]
    private async Task<ServiceResponse> DispatchSheetAsync(string actionString, ServiceRequest request)

    {

        if (ServiceRegistry.Sheet.TryParseAction(actionString, out var sheetAction))

        {

            // CopyToFile/MoveToFile are atomic operations without session

            if (sheetAction is SheetAction.CopyToFile or SheetAction.MoveToFile)

            {
                using var temporary = _sessionManager.OwnTemporaryBatches(request.Command, request.CallerCancellation);

                try

                {

                    var result = ServiceRegistry.Sheet.DispatchToCore(
                        _sheetCommands, sheetAction, null!, request.Args);
                    temporary.Complete();
                    return WrapResult(result);

                }

                catch (Exception ex)

                {
                    Exception? recoveryFailure = null;
                    try
                    {
                        temporary.RetainFailure();
                    }
                    catch (Exception recoveryError)
                    {
                        recoveryFailure = recoveryError;
                    }
                    if (temporary.SessionId is { } sessionId)
                        _knownSessionIds.TryAdd(sessionId, 0);
                    return CreateErrorResponse(ex, request.Command, temporary.SessionId,
                        temporary.SessionId is not null ? temporary.FilePaths : null, recoveryFailure);

                }

            }



            return await WithSessionAsync(request, batch =>

                WrapResult(ServiceRegistry.Sheet.DispatchToCore(_sheetCommands, sheetAction, batch, request.Args)));

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
                return WrapResult(ServiceRegistry.Range.DispatchToCore(_rangeCommands, ra, batch, request.Args));

            if (ServiceRegistry.RangeEdit.TryParseAction(actionString, out var rea))
                return WrapResult(ServiceRegistry.RangeEdit.DispatchToCore(_rangeCommands, rea, batch, request.Args));

            if (ServiceRegistry.RangeFormat.TryParseAction(actionString, out var rfa))
                return WrapResult(ServiceRegistry.RangeFormat.DispatchToCore(_rangeCommands, rfa, batch, request.Args));

            if (ServiceRegistry.RangeLink.TryParseAction(actionString, out var rla))
                return WrapResult(ServiceRegistry.RangeLink.DispatchToCore(_rangeCommands, rla, batch, request.Args));

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

                return WrapResult(ServiceRegistry.Table.DispatchToCore(_tableCommands, ta, batch, request.Args));

            if (ServiceRegistry.TableColumn.TryParseAction(actionString, out var tca))

                return WrapResult(ServiceRegistry.TableColumn.DispatchToCore(_tableCommands, tca, batch, request.Args));

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
        }, allowTimedOutOperation: windowAction is WindowAction.Show or WindowAction.Hide);
    }

    internal Action? BeforeSaveAsPublicationHookForTests { get; set; }
    internal Action? BeforeSaveAsReservationHookForTests { get; set; }

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
            if (workbookAction == WorkbookAction.SaveAs &&
                !string.IsNullOrWhiteSpace(request.SessionId))
            {
                if (_sessionManager.GetOperationStatus(request.SessionId)?.ExecutionPending == true)
                {
                    return new ServiceResponse
                    {
                        Success = false,
                        SessionId = request.SessionId,
                        ErrorCategory = "Busy",
                        ErrorMessage = "Cannot save-as: Excel work is still pending. The workbook remains open."
                    };
                }
                BeforeSaveAsReservationHookForTests?.Invoke();
                reservedPath = _sessionManager.ReserveSessionFilePath(
                    request.SessionId,
                    GetRequiredStringArgument(request.Args, "targetPath"));
            }

            try
            {
                var result = WrapResult(
                    ServiceRegistry.Workbook.DispatchToCore(_workbookCommands, workbookAction, batch, request.Args));

                if (reservedPath is not null)
                    BeforeSaveAsPublicationHookForTests?.Invoke();
                return result;
            }
            catch (Exception ex) when (ex is TimeoutException or OperationCanceledException)
            {
                if (reservedPath is not null && batch.HasTimedOutOperation)
                {
                    _sessionManager.RetainPendingSaveAs(request.SessionId!);
                    releaseReservation = false;
                }
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


    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Session operation failure-transport boundary returns explicit failed responses with the original exception and session context; confirmed fatal disconnects additionally remove dead-session tracking.")]
    private Task<ServiceResponse> WithSessionAsync(
        ServiceRequest request,
        Func<IExcelBatch, ServiceResponse> action,
        bool allowTimedOutOperation = false)
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

        var sessionError = TryBeginUsableSession(sessionId, out var batch, allowTimedOutOperation);
        if (sessionError != null)
        {
            return Task.FromResult(sessionError);
        }

        try
        {
            _sessionManager.SetOperationContext(sessionId, request.Command, request.CallerCancellation);
            var owner = batch is AttachedWorkbookBatch attached ? attached.Owner : batch as ExcelBatch;
            if (!string.Equals(request.Command, "session.notices", StringComparison.OrdinalIgnoreCase))
            {
                owner?.ThrowIfWorkbookUnavailable();
                if (!allowTimedOutOperation
                    && owner?.GetExecutionPreflightState() == WorkbookRefreshState.DialogOpen)
                    ExcelBusyException.ThrowIfNotReady(WorkbookRefreshState.DialogOpen, request.Command,
                        batch as IExcelBatchWindowControl);
            }
            if (string.Equals(request.Command, ServiceRegistry.Connection.GetRefreshStatusCommand, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(request.Command, ServiceRegistry.Connection.CancelRefreshCommand, StringComparison.OrdinalIgnoreCase))
            {
                ConnectionCommands.ValidateRefreshControlReadiness(batch!);
            }
            if (string.Equals(request.Command, ServiceRegistry.Connection.GetAccountSettingsCommand, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(request.Command, ServiceRegistry.Connection.SetAccountSettingsCommand, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(request.Command, ServiceRegistry.Connection.ClearAccountHintCommand, StringComparison.OrdinalIgnoreCase))
            {
                ConnectionCommands.ValidateAccountSettingsReadiness(batch!);
            }
            ServiceRegistry.ValidateWorkbookWriteAccess(request.Command, batch!, request.Args);
            var response = action(batch!);
            return Task.FromResult(response);
        }
        catch (TimeoutException ex) when (!IsFatalExcelDisconnect(ex))
        {
            return Task.FromResult(CreateErrorResponse(ex, request.Command, sessionId));
        }
        catch (OperationCanceledException ex) when (!IsFatalExcelDisconnect(ex))
        {
            return Task.FromResult(CreateErrorResponse(ex, request.Command, sessionId));
        }
        catch (Exception ex) when (IsFatalExcelDisconnect(ex))
        {
            return Task.FromResult(HandleExcelDisconnect(sessionId, ex));
        }
        catch (Exception ex)
        {
            // Check if Excel died with a non-COM exception — clean up dead session
            if (batch != null && !batch.IsExcelProcessAlive())
            {
                CleanupDeadSession(sessionId);
                return Task.FromResult(CreateExcelDisconnectedResponse(sessionId, ex,
                    $"Excel process for session '{sessionId}' is no longer running. Session has been cleaned up. Please reopen the file with a new session."));
            }

            return Task.FromResult(CreateErrorResponse(ex, request.Command, sessionId));
        }
        finally
        {
            _sessionManager.EndOperation(sessionId);
        }
    }

    private ServiceResponse HandleExcelDisconnect(string sessionId, Exception exception)
    {
        if (_sessionManager.GetSession(sessionId) is { } batch && batch.IsExcelProcessAlive())
        {
            _sessionManager.RevealForRecovery(sessionId);
            return new ServiceResponse
            {
                Success = false,
                SessionId = sessionId,
                ErrorCategory = "ExcelDisconnected",
                ErrorMessage = "Excel disconnected from automation, but its process is still live. " +
                    "The session and workbook remain owned; the operation's outcome is not confirmed. " +
                    exception.Message,
                ExceptionType = exception.GetType().Name,
                HResult = TryGetFatalComHResult(exception) is { } hresult ? $"0x{hresult:X8}" : null,
                InnerError = exception.InnerException?.Message
            };
        }
        CleanupDeadSession(sessionId);
        return CreateExcelDisconnectedResponse(sessionId, exception,
            "Excel disconnected and is no longer running. Dead-session tracking has been removed; reopen the workbook and verify whether changes were saved.");
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Cleanup follows an already-confirmed dead Excel session; secondary cleanup failure is diagnosed without replacing the original operation/disconnection error returned to the caller.")]
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

    private (int ProcessId, DateTime StartedUtc)? ResolveSessionProcess(string sessionId)
    {
        return _sessionManager.GetSessionProcessIdentity(sessionId) is { } identity
            ? (identity.ProcessId, DateTime.FromFileTimeUtc(identity.StartedAtUtcFileTime))
            : null;
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

            if (current is AggregateException aggregate)
            {
                foreach (var cause in aggregate.InnerExceptions)
                {
                    if (TryGetFatalComHResult(cause) is { } hresult)
                    {
                        return hresult;
                    }
                }

                return null;
            }
        }

        return null;
    }

    private static bool IsFatalComHResult(int hresult) =>
        hresult == ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE ||
        hresult == ResiliencePipelines.RPC_E_CALL_FAILED ||
        hresult == ResiliencePipelines.RPC_E_DISCONNECTED;

    internal static ServiceResponse AttachRequestContext(ServiceRequest request, ServiceResponse response)
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
            PartialState = response.PartialState,
            Result = response.Result
        };
    }

    private ServiceResponse CreateErrorResponse(
        Exception ex, string? command = null, string? sessionId = null,
        IReadOnlyList<string>? retainedFiles = null, Exception? recoveryFailure = null)
    {
        if (ex is SessionStartupPendingException pending)
        {
            sessionId = pending.SessionId;
            _knownSessionIds.TryAdd(sessionId, 0);
        }
        else if (ex is SessionRetainedException retained)
        {
            sessionId = retained.SessionId;
            _knownSessionIds.TryAdd(sessionId, 0);
        }
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
        if (recoveryFailure is not null)
        {
            innerError = $"{innerError}{(innerError is null ? "" : Environment.NewLine)}" +
                $"Recovery reveal also failed: {recoveryFailure.GetType().Name}: {recoveryFailure.Message}";
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
                InnerError = innerError,
                PartialState = SerializePartialState(pqEx.PartialState,
                    sessionId is not null ? _sessionManager.GetOperationStatus(sessionId) : null,
                    sessionId, retainedFiles)
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
                InnerError = innerError,
                PartialState = SerializePartialState(null,
                    sessionId is not null ? _sessionManager.GetOperationStatus(sessionId) : null,
                    sessionId, retainedFiles)
            }
        };
    }

    private static string? SerializePartialState(
        PowerQueryPartialState? query, ExcelOperationStatus? operation,
        string? sessionId = null, IReadOnlyList<string>? retainedFiles = null)
    {
        var state = query is null
            ? null
            : JsonSerializer.SerializeToNode(query, ServiceProtocol.JsonOptions)!.AsObject();
        if (retainedFiles is not null)
        {
            state ??= new System.Text.Json.Nodes.JsonObject();
            state["sessionId"] = sessionId;
            state["filePaths"] = JsonSerializer.SerializeToNode(retainedFiles, ServiceProtocol.JsonOptions);
            state["retained"] = true;
        }
        if (operation is { WaitExpired: true } || operation is not null && retainedFiles is not null)
        {
            state ??= new System.Text.Json.Nodes.JsonObject();
            state["operation"] = JsonSerializer.SerializeToNode(operation, ServiceProtocol.JsonOptions);
        }
        return state?.ToJsonString(ServiceProtocol.JsonOptions);
    }

    private ServiceResponse? TryBeginUsableSession(
        string sessionId, out IExcelBatch? batch, bool allowTimedOutOperation = false)
    {
        if (!_sessionManager.TryBeginOperation(
            sessionId,
            out batch,
            out var errorMessage,
            out var sessionError,
            allowTimedOutOperation))
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
                    SessionOperationError.TimedOutOrCancelled => "SessionUnavailable",
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
        lock (_shutdownLock)
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                ShutDownSessions();
            }
            finally
            {
                _daemonHost.RequestShutdown();
                _daemonHost.Dispose();
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
public sealed class SessionTestArgs
{
    public string? FilePath { get; set; }
    public int? TimeoutSeconds { get; set; }
}
