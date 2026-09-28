using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Formatting;
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
using Sbroenne.ExcelMcp.Service.Mac;
using StreamJsonRpc;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Commands.Window;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Sbroenne.ExcelMcp.Core.Commands.XmlMap;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
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
    private readonly MacPowerQueryHelperDispatcher? _macPowerQueryHelperDispatcher;
    private readonly MacVbaHelperDispatcher? _macVbaHelperDispatcher;
    private readonly Func<string, TimeSpan, Task<JsonElement>>? _getMacHelperCapabilities;
    private readonly IReadOnlySet<string> _macPowerQueryCandidateActions =
        new HashSet<string>(StringComparer.Ordinal);
    private readonly IReadOnlySet<string> _macVbaCandidateActions =
        new HashSet<string>(StringComparer.Ordinal);
    private readonly MacVbaPreflightResult _macVbaPreflight =
        new(
            MacMacroExecutionAvailability.Unknown,
            MacVbaProjectModelAccess.Unknown);
    private readonly ConcurrentDictionary<string, byte> _knownSessionIds = new(StringComparer.Ordinal);
    private readonly DaemonHost _daemonHost;
    private readonly DateTime _startTime = DateTime.UtcNow;
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
    private readonly QueryTableCommands _queryTableCommands = new();
    private readonly NamedRangeCommands _namedRangeCommands = new();
    private readonly ConditionalFormattingCommands _conditionalFormatCommands = new();
    private readonly VbaCommands _vbaCommands = new();
    private readonly DataModelCommands _dataModelCommands = new();
    private readonly CalculationModeCommands _calculationModeCommands = new();
    private readonly ScreenshotCommands _screenshotCommands = new();
    private readonly DiagCommands _diagCommands = new();
    private readonly DrawingCommands _drawingCommands = new();
    private readonly WindowCommands _windowCommands = new();
    private readonly WorkbookCommands _workbookCommands = new();
    private readonly PythonInExcelCommands _pythonInExcelCommands = new();
    private readonly AnalysisCommands _analysisCommands = new();
    private readonly XmlMapCommands _xmlMapCommands = new();
    private readonly FileCommands _fileCommands = new();

    public ExcelMcpService()
    {
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
        _macPowerQueryHelperDispatcher = null;
        _getMacHelperCapabilities = null;
        if (OperatingSystem.IsMacOS())
        {
            _macBackend = new MacExcelBackend();
            _macSessionManager = new MacExcelSessionManager(_macBackend);
            var helperClient = new MacVbaHelperClient(_macBackend);
            _getMacHelperCapabilities = helperClient.GetCapabilitiesAsync;
            _macPowerQueryHelperDispatcher = new MacPowerQueryHelperDispatcher(
                helperClient.DispatchAsync);
            _macPowerQueryCandidateActions =
                MacPowerQueryHelperCapabilities.GetExplicitOptIn();
            _macVbaHelperDispatcher = new MacVbaHelperDispatcher(
                helperClient.DispatchAsync);
            _macVbaCandidateActions = MacVbaHelperCapabilities.GetExplicitOptIn();
            _macVbaPreflight = MacVbaPreflight.Check();
        }
    }

    internal ExcelMcpService(
        MacExcelBackend macBackend,
        Func<string, TimeSpan, Task<JsonElement>> getMacHelperCapabilities,
        MacPowerQueryHelperDispatch macPowerQueryHelperDispatch,
        IReadOnlySet<string>? macPowerQueryCandidateActions = null,
        IReadOnlySet<string>? macVbaCandidateActions = null,
        MacVbaPreflightResult? macVbaPreflight = null)
    {
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
        _macBackend = macBackend;
        _macSessionManager = new MacExcelSessionManager(macBackend);
        _getMacHelperCapabilities = getMacHelperCapabilities;
        _macPowerQueryHelperDispatcher = new MacPowerQueryHelperDispatcher(
            macPowerQueryHelperDispatch);
        _macPowerQueryCandidateActions = macPowerQueryCandidateActions
            ?? new HashSet<string>(StringComparer.Ordinal);
        _macVbaHelperDispatcher = new MacVbaHelperDispatcher(
            macPowerQueryHelperDispatch);
        _macVbaCandidateActions = macVbaCandidateActions
            ?? new HashSet<string>(StringComparer.Ordinal);
        _macVbaPreflight = macVbaPreflight
            ?? new(
                MacMacroExecutionAvailability.Unknown,
                MacVbaProjectModelAccess.Unknown);
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

            if (_macSessionManager != null)
            {
                var macResponse = category == "service"
                    ? HandleServiceCommand(action)
                    : category == "session"
                        ? await HandleMacSessionCommandAsync(action, request)
                        : await DispatchMacCommandAsync(category, action, request);
                return AttachRequestContext(request, macResponse);
            }

            ServiceResponse response = category switch
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
                "querytable" => await DispatchSimpleAsync<QueryTableAction>(action, request,
                    ServiceRegistry.QueryTable.TryParseAction,
                    (a, batch) => ServiceRegistry.QueryTable.DispatchToCore(_queryTableCommands, a, batch, request.Args)),
                "calculation" => await DispatchSimpleAsync<CalculationAction>(action, request,
                    ServiceRegistry.Calculation.TryParseAction,
                    (a, batch) => ServiceRegistry.Calculation.DispatchToCore(_calculationModeCommands, a, batch, request.Args)),
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
                "workbook" => await DispatchWorkbookAsync(action, request),
                "diag" => DispatchSessionless(action, request),
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
        catch (MacExcelOperationException ex)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = ex.ErrorCategory,
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name,
                Command = request.Command,
                SessionId = request.SessionId
            };
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
                ErrorMessage = $"Unknown service action: {action}"
            }
        };
    }

    private ServiceResponse HandleShutdown()
    {
        _daemonHost.RequestShutdownAfterResponse();
        return new ServiceResponse { Success = true };
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
                canClose = session.PendingOperations == 0 && !session.IsClosing,
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
            return HandleSessionTest(request);
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
            var session = _macSessionManager!.Sessions.FirstOrDefault(
                candidate => candidate.SessionId == request.SessionId);
            var closed = await _macSessionManager!.CloseAsync(request.SessionId, closeArgs?.Save ?? false);
            if (closed)
            {
                if (session is not null)
                {
                    await TryUnregisterOfficeSessionAsync(session);
                }
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
        if (action == "create"
            && args.MacroEnabled.HasValue
            && args.MacroEnabled.Value != macroEnabled)
        {
            throw new ArgumentException(
                $"macroEnabled must be {macroEnabled.ToString().ToLowerInvariant()} for a '{extension}' workbook.");
        }

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

    private async Task<ServiceResponse> DispatchMacCommandAsync(
        string category,
        string action,
        ServiceRequest request)
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

        var command = $"{category}.{action}";
        var officeCandidateEnabled = MacOfficeActionCatalog.TryGet(command, out _)
            && MacOfficeBridgeConfiguration.IsActionEnabled(command);
        var capability = MacCommandCapabilities.Get(
            command,
            officeCandidateEnabled);
        var scenarioAcceptance = CanUseScenarioForAcceptance(
            command,
            Environment.GetEnvironmentVariable("EXCELMCP_MAC_SCENARIO_E2E"),
            MacVbaHelperClient.GetInstallation());
        // Query variants are gated by helper proof before dispatch, not static inventory availability.
        var powerQueryRouteSelection = category == "powerquery"
            && ServiceRegistry.PowerQuery.TryParseAction(action, out _);
        var vbaRouteSelection = category == "vba"
            && ServiceRegistry.Vba.TryParseAction(action, out _);
        if (!capability.IsAvailable
            && !scenarioAcceptance
            && !powerQueryRouteSelection
            && !vbaRouteSelection)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "PlatformNotSupported",
                ErrorMessage = capability.UnavailableMessage
            };
        }

        try
        {
            return await _macSessionManager!.ExecuteAsync(request.SessionId, async session =>
            {
                var arguments = string.IsNullOrWhiteSpace(request.Args)
                    ? new JsonObject()
                    : JsonNode.Parse(request.Args)?.AsObject() ?? new JsonObject();
                ResolveMacFileArguments(category, action, arguments);

                if (capability.RequiredTier == MacCapabilityTier.OfficeAddIn)
                {
                    if (!MacOfficeActionCatalog.TryGet(command, out var officeAction))
                    {
                        throw new InvalidOperationException(
                            $"Office.js metadata is missing for '{command}'.");
                    }
                    try
                    {
                        using var officeClient = MacOfficeBridgeClient.CreateDefault();
                        var officeResult = await officeClient.InvokeAsync(
                            session.SessionId,
                            session.FilePath,
                            command,
                            arguments,
                            session.OperationTimeout,
                            officeAction.Mutation);
                        return new ServiceResponse
                        {
                            Success = true,
                            Result = officeResult.GetRawText()
                        };
                    }
                    catch (MacOfficeMutationUncertainException ex)
                    {
                        session.MarkUnsafe(ex.Message);
                        throw;
                    }
                }

                if (category == "powerquery")
                {
                    return await DispatchMacPowerQueryAsync(action, request, session);
                }
                if (category == "vba")
                {
                    return await DispatchMacVbaAsync(action, arguments, session);
                }

                var arguments = string.IsNullOrWhiteSpace(request.Args)
                    ? new JsonObject()
                    : JsonNode.Parse(request.Args)?.AsObject() ?? new JsonObject();
                if (category == "pythoninexcel")
                {
                    MacPythonInExcelArguments.Prepare(action, arguments, session.OperationTimeout);
                }
                arguments["filePath"] = session.FilePath;
                MacRangeArguments.Prepare(category, action, arguments);
                ValidateMacRangeFormatArguments(category, action, arguments);
                var result = await _macBackend!.InvokeAsync(
                    command,
                    arguments,
                    session.OperationTimeout,
                    allowFailureResult: category == "pythoninexcel");
                return new ServiceResponse
                {
                    Success = true,
                    Result = result.GetRawText()
                };
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
        catch (MacVbaHelperException ex)
        {
            if (string.Equals(ex.Category, "RecoveryRequired", StringComparison.Ordinal)
                || string.Equals(ex.Code, "rollback_failed", StringComparison.Ordinal))
            {
                await InvalidateMacSessionAsync(request.SessionId);
            }
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = ex.Category,
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name
            };
        }
        catch (TimeoutException ex)
        {
            await InvalidateMacSessionAsync(request.SessionId);
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Timeout",
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name
            };
        }
        catch (MacOfficeBridgeException ex)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = ex.ErrorCategory,
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name
            };
        }
        catch (MacExcelOperationException ex)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = ex.ErrorCategory,
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name
            };
        }
        catch (MacVbaHelperException ex)
        {
            return new ServiceResponse
            {
                Success = false,
                ErrorCategory = ex.Category,
                ErrorMessage = ex.Message,
                ExceptionType = ex.GetType().Name
            };
        }
        catch (Exception ex)
        {
            return CreateErrorResponse(ex);
        }
    }

    private async Task InvalidateMacSessionAsync(string sessionId)
    {
        _macSessionManager!.RequireRecovery(sessionId);
        try
        {
            await _macSessionManager.CloseAsync(sessionId, save: false);
        }
        catch
        {
            // Shared Excel is never killed; recovery gating remains if exact close did not complete.
        }
    }

    private async Task<ServiceResponse> DispatchMacPowerQueryAsync(
        string action,
        ServiceRequest request,
        MacExcelSession session)
    {
        var arguments = string.IsNullOrWhiteSpace(request.Args)
            ? new JsonObject()
            : JsonNode.Parse(request.Args)?.AsObject()
                ?? throw new ArgumentException($"Power Query {action} arguments must be an object.");
        NormalizeMacPowerQueryArguments(action, arguments);
        var route = MacPowerQueryRouteSelector.Select(
            action,
            arguments,
            new HashSet<string>(StringComparer.Ordinal));
        TimeSpan? helperTimeout = null;
        long helperStartedAt = 0;
        if (route.Kind == MacPowerQueryRouteKind.Unsupported
            && _macPowerQueryHelperDispatcher is not null
            && _getMacHelperCapabilities is not null)
        {
            helperTimeout = GetMacPowerQueryTimeout(
                action,
                arguments,
                session.OperationTimeout);
            helperStartedAt = Stopwatch.GetTimestamp();
            var capabilities = await _getMacHelperCapabilities(
                session.FilePath,
                helperTimeout.Value);
            route = MacPowerQueryRouteSelector.Select(
                action,
                arguments,
                MacPowerQueryHelperCapabilities.Parse(
                    capabilities,
                    _macPowerQueryCandidateActions));
        }
        if (route.Kind == MacPowerQueryRouteKind.Unsupported)
        {
            throw UnsupportedMacPowerQueryVariant(
                $"powerquery.{action}",
                route.UnavailableReason ?? "the selected helper method is unavailable");
        }
        await FormatMacPowerQueryMCodeAsync(action, arguments, route);
        if (route.Kind == MacPowerQueryRouteKind.Helper)
        {
            var result = await _macPowerQueryHelperDispatcher!.DispatchAsync(
                route,
                session.FilePath,
                RemainingMacPowerQueryTimeout(
                    helperStartedAt,
                    helperTimeout
                        ?? GetMacPowerQueryTimeout(
                            action,
                            arguments,
                            session.OperationTimeout)),
                action,
                arguments);
            return new ServiceResponse
            {
                Success = true,
                Result = result.GetRawText()
            };
        }

        var state = await _macBackend!.InvokeAsync(
            "workbook.state",
            new { filePath = session.FilePath },
            session.OperationTimeout);
        if (!state.GetProperty("saved").GetBoolean())
        {
            throw new InvalidOperationException(
                "Power Query package reads on macOS require a saved workbook. " +
                "Save or discard the current workbook changes, then retry.");
        }

        var queries = MacPowerQueryPackage.ReadQueries(session.FilePath);
        var loads = MacPowerQueryPackage.ReadWorksheetLoads(session.FilePath);

        PowerQueryInfo CreateInfo(MacPowerQueryDefinition query)
        {
            var matchingLoads = loads.Where(candidate =>
                PowerQueryHelpers.MatchesMashupLocation(candidate.Connection, query.Name)).ToArray();
            if (matchingLoads.Length > 1)
            {
                throw new InvalidDataException(
                    $"Power Query '{query.Name}' has multiple worksheet destinations. " +
                    "macOS cannot report or refresh this state safely.");
            }
            var load = matchingLoads.SingleOrDefault();
            var isConnectionOnly = load is null;
            return new PowerQueryInfo
            {
                Name = query.Name,
#pragma warning disable CS0618
                Formula = query.Formula,
#pragma warning restore CS0618
                FormulaPreview = query.Formula.Length > 80
                    ? query.Formula[..77] + "..."
                    : query.Formula,
                CharacterCount = query.Formula.Length,
                LoadMode = isConnectionOnly
                    ? PowerQueryLoadMode.ConnectionOnly
                    : PowerQueryLoadMode.LoadToTable,
                TargetSheet = load?.SheetName,
                IsConnectionOnly = isConnectionOnly,
                IsLoadedToDataModel = false
            };
        }

        if (action == "list")
        {
            var result = new PowerQueryListResult
            {
                Success = true,
                FilePath = session.FilePath,
                Queries = queries.Select(CreateInfo).ToList()
            };
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions)
            };
        }

        var queryName = arguments["queryName"]?.GetValue<string>()
            ?? throw new ArgumentException("queryName is required.");
        var query = queries.SingleOrDefault(candidate =>
                string.Equals(candidate.Name, queryName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Query '{queryName}' not found.");
        var info = CreateInfo(query);

        if (action == "get-load-config")
        {
            return SerializeMacResult(new PowerQueryLoadConfigResult
            {
                Success = true,
                FilePath = session.FilePath,
                QueryName = query.Name,
                HasConnection = !info.IsConnectionOnly,
                LoadMode = info.LoadMode,
                TargetSheet = info.TargetSheet,
                IsLoadedToDataModel = false
            });
        }

        if (action == "update")
        {
            var mCode = arguments["mCode"]!.GetValue<string>();
            var refresh = arguments["refresh"]?.GetValue<bool>() ?? true;
            if (refresh)
            {
                throw UnsupportedMacPowerQueryVariant(
                    "powerquery.update",
                    "refresh was requested but exact refresh completion and error propagation " +
                    "are not yet proven through the production Mac bridge");
            }

            await _macSessionManager!.MutatePackageAsync(
                session,
                workingPath =>
                {
                    MacPowerQueryPackage.UpdateQuery(workingPath, query.Name, mCode);
                    var updated = MacPowerQueryPackage.ReadQueries(workingPath)
                        .Single(candidate =>
                            string.Equals(candidate.Name, query.Name, StringComparison.OrdinalIgnoreCase));
                    if (!string.Equals(updated.Formula, mCode, StringComparison.Ordinal))
                    {
                        throw new InvalidDataException(
                            $"Power Query '{query.Name}' did not retain the requested M code.");
                    }
                    _ = MacPowerQueryPackage.ReadWorksheetLoads(workingPath);
                },
                static () => Task.CompletedTask);
            return SerializeMacResult(new OperationResult
            {
                Success = true,
                FilePath = session.FilePath
            });
        }

        if (action != "view")
        {
            throw new InvalidOperationException($"Unhandled macOS Power Query action '{action}'.");
        }

        var view = new PowerQueryViewResult
        {
            Success = true,
            FilePath = session.FilePath,
            QueryName = query.Name,
            MCode = query.Formula,
            CharacterCount = query.Formula.Length,
            LoadMode = info.LoadMode,
            TargetSheet = info.TargetSheet,
            HasConnection = !info.IsConnectionOnly,
            IsLoadedToDataModel = false,
            IsConnectionOnly = info.IsConnectionOnly
        };
        return SerializeMacResult(view);
    }

    private async Task<ServiceResponse> DispatchMacVbaAsync(
        string action,
        JsonObject arguments,
        MacExcelSession session)
    {
        var isMacroWorkbook = string.Equals(
            Path.GetExtension(session.FilePath),
            ".xlsm",
            StringComparison.OrdinalIgnoreCase);
        if (!isMacroWorkbook)
        {
            if (action == "list")
            {
                return new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(
                        new
                        {
                            success = true,
                            filePath = session.FilePath,
                            scripts = Array.Empty<object>()
                        },
                        ServiceProtocol.JsonOptions)
                };
            }
            throw new ArgumentException(
                "VBA operations require a macro-enabled workbook (.xlsm).");
        }

        var initialRoute = MacVbaRouteSelector.Select(
            action,
            arguments,
            capabilities: null,
            _macVbaPreflight);
        if (!initialRoute.IsAvailable)
        {
            throw UnsupportedMacVbaVariant(
                $"vba.{action}",
                initialRoute.UnavailableReason
                    ?? "the selected helper method is unavailable");
        }

        var timeout = GetMacVbaTimeout(action, arguments, session.OperationTimeout);
        var startedAt = Stopwatch.GetTimestamp();
        var capabilityResult = await _getMacHelperCapabilities!(
            session.FilePath,
            timeout);
        var route = MacVbaRouteSelector.Select(
            action,
            arguments,
            MacVbaHelperCapabilities.Parse(
                capabilityResult,
                _macVbaCandidateActions),
            _macVbaPreflight);
        if (!route.IsAvailable)
        {
            throw UnsupportedMacVbaVariant(
                $"vba.{action}",
                route.UnavailableReason
                    ?? "the selected helper method is unavailable");
        }

        var result = await _macVbaHelperDispatcher!.DispatchAsync(
            route,
            session.FilePath,
            RemainingMacHelperTimeout(startedAt, timeout, "VBA"),
            action);
        return new ServiceResponse
        {
            Success = true,
            Result = result.GetRawText()
        };
    }

    private static void NormalizeMacPowerQueryArguments(
        string action,
        JsonObject arguments)
    {
        if (action is not ("create" or "update" or "evaluate"))
        {
            return;
        }

        var mCode = ParameterTransforms.ResolveFileOrValue(
            arguments["mCode"]?.GetValue<string>(),
            arguments["mCodeFile"]?.GetValue<string>(),
            "mCode");
        if (string.IsNullOrWhiteSpace(mCode))
        {
            throw new ArgumentException(
                action == "evaluate"
                    ? "M code is required for evaluate action."
                    : "M code cannot be empty.");
        }
        arguments["mCode"] = mCode;
        arguments.Remove("mCodeFile");
    }

    private static async Task FormatMacPowerQueryMCodeAsync(
        string action,
        JsonObject arguments,
        MacPowerQueryRoute route)
    {
        if (action is not ("create" or "update")
            || arguments["formatMCode"]?.GetValue<bool>() != true)
        {
            return;
        }

        var formattedMCode = await MCodeFormatter.FormatAsync(
            arguments["mCode"]!.GetValue<string>());
        arguments["mCode"] = formattedMCode;
        if (route.Kind == MacPowerQueryRouteKind.Helper)
        {
            route.HelperArguments!["formula"] = formattedMCode;
        }
    }

    private static TimeSpan GetMacPowerQueryTimeout(
        string action,
        JsonObject arguments,
        TimeSpan sessionTimeout)
    {
        if (action is "refresh" or "refresh-all")
        {
            var seconds = arguments["timeout"]?.GetValue<double>() ?? 0;
            return seconds > 0
                ? TimeSpan.FromSeconds(seconds)
                : ComInteropConstants.DataOperationTimeout;
        }
        return action is "create" or "load-to" or "evaluate"
            ? ComInteropConstants.DataOperationTimeout
            : sessionTimeout;
    }

    private static TimeSpan RemainingMacPowerQueryTimeout(
        long startedAt,
        TimeSpan timeout)
        => RemainingMacHelperTimeout(startedAt, timeout, "Power Query");

    private static TimeSpan RemainingMacHelperTimeout(
        long startedAt,
        TimeSpan timeout,
        string feature)
    {
        if (timeout == Timeout.InfiniteTimeSpan)
        {
            return timeout;
        }

        var remaining = timeout - Stopwatch.GetElapsedTime(startedAt);
        if (remaining <= TimeSpan.Zero)
        {
            throw new TimeoutException(
                $"The macOS {feature} helper operation exceeded its shared " +
                "capability-and-dispatch deadline. The workbook session is no longer safe to use.");
        }
        return remaining;
    }

    private static TimeSpan GetMacVbaTimeout(
        string action,
        JsonObject arguments,
        TimeSpan sessionTimeout)
    {
        if (action != "run" || arguments["timeout"] is null)
        {
            return sessionTimeout;
        }
        if (arguments["timeout"] is not JsonValue value
            || !value.TryGetValue<double>(out var seconds)
            || !double.IsInteger(seconds)
            || seconds <= 0
            || seconds > int.MaxValue / 1000d)
        {
            throw new ArgumentException(
                "timeout must be a positive whole number of seconds.");
        }
        return TimeSpan.FromSeconds(seconds);
    }

    private static MacExcelOperationException UnsupportedMacPowerQueryVariant(
        string command,
        string reason) =>
        new(
            "PlatformNotSupported",
            $"Command '{command}' cannot run on macOS because {reason}. " +
            "Only saved-package inspection and package-only update with refresh=false " +
            "are supported in this release; " +
            "the complete action remains available on Windows.");

    private static MacExcelOperationException UnsupportedMacVbaVariant(
        string command,
        string reason) =>
        new(
            "PlatformNotSupported",
            $"Command '{command}' cannot run on macOS because {reason}. " +
            "The optional helper remains disabled by default until exact public " +
            "CLI and MCP acceptance has proven this action.");

    private static ServiceResponse SerializeMacResult(ResultBase result) =>
        new()
        {
            Success = true,
            Result = JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions)
        };

    private static void ResolveMacFileArguments(
        string category,
        string action,
        JsonObject arguments)
    {
        if (category == "table" && action == "append")
        {
            var rows = arguments["rows"]?.Deserialize<List<List<object?>>>(ServiceProtocol.JsonOptions);
            var rowsFile = arguments["rowsFile"]?.GetValue<string>();
            arguments["rows"] = JsonSerializer.SerializeToNode(
                ParameterTransforms.ResolveValuesOrFile(rows, rowsFile, "rows"),
                ServiceProtocol.JsonOptions);
            arguments.Remove("rowsFile");
            return;
        }

        if (category == "vba" && action is "import" or "update")
        {
            arguments["source"] = ParameterTransforms.ResolveFileOrValue(
                arguments["vbaCode"]?.GetValue<string>(),
                arguments["vbaCodeFile"]?.GetValue<string>(),
                "vbaCode");
            arguments.Remove("vbaCode");
            arguments.Remove("vbaCodeFile");
            return;
        }

        if (category != "range")
        {
            ValidateMacRangeFormatArguments(category, action, arguments);
            return;
        }

        MacRangeArguments.Prepare(category, action, arguments);
    }

    private static void ValidateMacRangeFormatArguments(
        string category,
        string action,
        JsonObject arguments)
    {
        if (category != "rangeformat")
        {
            return;
        }

        if (action == "set-column-width")
        {
            var width = arguments["columnWidth"]?.GetValue<double>()
                ?? throw new ArgumentException("columnWidth is required.");
            if (width is < 0.25 or > 409)
            {
                throw new ArgumentException("columnWidth must be between 0.25 and 409 points");
            }
        }
        else if (action == "set-row-height")
        {
            var height = arguments["rowHeight"]?.GetValue<double>()
                ?? throw new ArgumentException("rowHeight is required.");
            if (height is < 0 or > 409)
            {
                throw new ArgumentException("rowHeight must be between 0 and 409 points");
            }
        }
    }

    private ServiceResponse HandleSessionCommand(string action, ServiceRequest request)
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

    private static void ValidateSessionActionArguments(string action, string? argsJson)
    {
        var allowedParameters = action switch
        {
            "create" => new HashSet<string>(
                ["filePath", "macroEnabled", "show", "timeoutSeconds"],
                StringComparer.Ordinal),
            "open" => new HashSet<string>(
                ["filePath", "show", "timeoutSeconds"],
                StringComparer.Ordinal),
            "close" => new HashSet<string>(["save"], StringComparer.Ordinal),
            "test" => new HashSet<string>(
                ["filePath", "timeoutSeconds"],
                StringComparer.Ordinal),
            _ => []
        };
        var unknownParameters = ServiceRegistry.GetJsonPropertyNames(argsJson, includeNullValues: true)
            .Where(parameter => !allowedParameters.Contains(parameter))
            .ToArray();
        if (unknownParameters.Length > 0)
        {
            throw new ArgumentException(
                $"Unknown parameter(s) for session.{action}: {string.Join(", ", unknownParameters)}.");
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
        var extensionIsMacroEnabled = string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase);
        if (args.MacroEnabled.HasValue && args.MacroEnabled.Value != extensionIsMacroEnabled)
        {
            throw new ArgumentException(
                $"macroEnabled must be {extensionIsMacroEnabled.ToString().ToLowerInvariant()} for a '{extension}' workbook.");
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
        var fullPath = FilePathValidation.NormalizeAbsoluteWindowsPath(args.FilePath);

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
                && result.Extension is ".xlsx" or ".xlsm"
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



        return await WithSessionAsync(request.SessionId, batch => WrapResult(dispatch(action, batch)));

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

                    return CreateErrorResponse(ex);

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



        return new ServiceResponse
        {
            Success = false,
            ErrorCategory = "InvalidInput",
            ErrorMessage = $"Unknown sheet action: {actionString}"
        };

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

        return await WithSessionAsync(request.SessionId, batch =>

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

        return await WithSessionAsync(request.SessionId, batch =>
        {
            string? reservedPath = null;
            var releaseReservation = true;
            if (workbookAction == WorkbookAction.SaveAs &&
                !string.IsNullOrWhiteSpace(request.SessionId))
            {
                reservedPath = _sessionManager.ReserveSessionFilePath(
                    request.SessionId,
                    GetRequiredStringArgument(request.Args, "targetPath"));
            }

            try
            {
                var result = WrapResult(
                    ServiceRegistry.Workbook.DispatchToCore(_workbookCommands, workbookAction, batch, request.Args));

                if (result.Success && reservedPath != null)
                {
                    _sessionManager.UpdateSessionFilePath(request.SessionId!, batch.WorkbookPath);
                }

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


    private Task<ServiceResponse> WithSessionAsync(string? sessionId, Func<IExcelBatch, ServiceResponse> action)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "InvalidInput",
                ErrorMessage = "sessionId is required"
            });
        }

        var sessionError = TryBeginUsableSession(sessionId, out var batch);
        if (sessionError != null)
        {
            return Task.FromResult(sessionError);
        }

        try
        {
            var response = action(batch!);
            return Task.FromResult(response);
        }
        catch (TimeoutException ex)
        {
            // Operation timed out — Excel COM call is hung (IDispatch.Invoke stuck).
            // Force-close the session to trigger the force-kill path in ExcelBatch.Dispose(),
            // which will kill the hung Excel process and release the STA thread.
            try
            {
                _sessionManager.CloseSession(sessionId, save: false, force: true);
            }
            catch (Exception cleanupEx)
            {
                System.Diagnostics.Debug.WriteLine($"Session cleanup failed for {sessionId}: {cleanupEx.Message}");
            }
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Timeout",
                ErrorMessage = $"Excel operation timed out and the session has been closed: {ex.Message} " +
                               "Please reopen the file with a new session.",
                ExceptionType = ex.GetType().Name
            });
        }
        catch (OperationCanceledException)
        {
            // Caller cancelled (e.g., VS Code cancelled the tool call) while a COM operation
            // may still be running on the STA thread. ExcelBatch.Execute sets _operationTimedOut
            // on cancellation, but nobody calls Dispose() — the session stays alive with a
            // stuck STA thread, and all subsequent requests queue up and hang.
            // Force-close the session to kill the hung Excel process and release the STA thread.
            try
            {
                _sessionManager.CloseSession(sessionId, save: false, force: true);
            }
            catch (Exception cleanupEx)
            {
                System.Diagnostics.Debug.WriteLine($"Session cleanup failed for {sessionId}: {cleanupEx.Message}");
            }
            return Task.FromResult(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "Cancelled",
                ErrorMessage = $"Operation was cancelled and the session has been closed. " +
                               "The Excel COM thread may have been unresponsive. " +
                               "Please reopen the file with a new session.",
                ExceptionType = nameof(OperationCanceledException)
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
        catch (InvalidOperationException ex) when (
            ex.Message.Contains("no longer running", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("process", StringComparison.OrdinalIgnoreCase))
        {
            // Excel process detected as dead before COM call (ExcelBatch pre-check)
            CleanupDeadSession(sessionId);
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
                CleanupDeadSession(sessionId);
                return Task.FromResult(CreateExcelDisconnectedResponse(sessionId, ex,
                    $"Excel process for session '{sessionId}' disconnected during the operation. Session has been cleaned up. Please reopen the file with a new session."));
            }

            // Check if Excel died with a non-COM exception — clean up dead session
            if (batch != null && !batch.IsExcelProcessAlive())
            {
                CleanupDeadSession(sessionId);
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

        _shutdownCts.Cancel();
        _macSessionManager?.Dispose();
        _sessionManager.Dispose();
        _shutdownCts.Dispose();
    }
}

// === ARGUMENT TYPES (Session only - all other args are now generated in ServiceRegistry) ===

// Session
public sealed class SessionOpenArgs
{
    public string? FilePath { get; set; }
    public bool? MacroEnabled { get; set; }
    public bool Show { get; set; }
    public int? TimeoutSeconds { get; set; }
}
public sealed class SessionCloseArgs { public bool Save { get; set; } }
public sealed class SessionTestArgs
{
    public string? FilePath { get; set; }
    public int? TimeoutSeconds { get; set; }
}
