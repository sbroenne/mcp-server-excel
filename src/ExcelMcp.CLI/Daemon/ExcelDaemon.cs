using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Models.Actions;

namespace Sbroenne.ExcelMcp.CLI.Daemon;

/// <summary>
/// The Excel daemon process. Holds SessionManager and executes Core commands.
/// Runs as a background process, accepting commands via named pipe.
/// Architecture mirrors MCP Server: CLI sends serialized requests → Daemon executes → Returns JSON.
/// </summary>
internal sealed class ExcelDaemon : IDisposable
{
    private readonly SessionManager _sessionManager = new();
    private readonly CancellationTokenSource _shutdownCts = new();
    private readonly TimeSpan _idleTimeout;
    private readonly DateTime _startTime = DateTime.UtcNow;
    private Mutex? _instanceMutex;
    private DaemonTray? _tray;
    private DateTime _lastActivityTime = DateTime.UtcNow;
    private bool _disposed;

    // Core command instances - use concrete types per CA1859
    private readonly RangeCommands _rangeCommands = new();
    private readonly SheetCommands _sheetCommands = new();
    private readonly TableCommands _tableCommands = new();
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly PivotTableCommands _pivotTableCommands = new();
    private readonly ChartCommands _chartCommands = new();
    private readonly ConnectionCommands _connectionCommands = new();
    private readonly NamedRangeCommands _namedRangeCommands = new();
    private readonly ConditionalFormattingCommands _conditionalFormatCommands = new();
    private readonly VbaCommands _vbaCommands = new();
    private readonly DataModelCommands _dataModelCommands = new();

    public ExcelDaemon(TimeSpan? idleTimeout = null)
    {
        _idleTimeout = idleTimeout ?? DefaultIdleTimeout;
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
    }

    public static readonly TimeSpan DefaultIdleTimeout = TimeSpan.FromMinutes(10);



    public DateTime StartTime => _startTime;
    public int SessionCount => _sessionManager.GetActiveSessions().Count;

    /// <summary>
    /// Runs the daemon, listening for commands on the named pipe.
    /// </summary>
    public async Task RunAsync()
    {
        // Acquire single-instance mutex
        _instanceMutex = DaemonSecurity.TryAcquireSingleInstanceMutex();
        if (_instanceMutex == null)
        {
            throw new InvalidOperationException("Another daemon instance is already running");
        }

        // Track that we own the lock file (only delete it if we created it)
        bool ownsLockFile = false;

        Thread? trayThread = null;

        try
        {
            // Write lock file - now we own it
            DaemonSecurity.WriteLockFile(Environment.ProcessId);
            ownsLockFile = true;

            // Start tray UI on STA thread (Windows Forms requires STA)
            trayThread = new Thread(() => RunTrayLoop())
            {
                IsBackground = true,
                Name = "DaemonTray"
            };
            trayThread.SetApartmentState(ApartmentState.STA);
            trayThread.Start();

            // Start idle monitor
            var idleMonitorTask = MonitorIdleTimeoutAsync(_shutdownCts.Token);

            // Main pipe server loop
            await RunPipeServerAsync(_shutdownCts.Token);

            await idleMonitorTask;
        }
        finally
        {
            // Cleanup tray
            if (_tray != null)
            {
                try
                {
                    // Signal the tray thread to exit
                    Application.Exit();
                }
                catch { }
            }

            // Only delete lock file if we created it
            if (ownsLockFile)
            {
                DaemonSecurity.DeleteLockFile();
            }

            // _instanceMutex is guaranteed non-null here (method throws if null)
            _instanceMutex.ReleaseMutex();
            _instanceMutex.Dispose();
            _instanceMutex = null;
        }
    }

    private void RunTrayLoop()
    {
        try
        {
            _tray = new DaemonTray(_sessionManager, RequestShutdown);

            // Run Windows Forms message loop - this blocks until Application.Exit() is called
            Application.Run();
        }
        finally
        {
            // Ensure _tray is disposed even if exception occurs
            _tray?.Dispose();
            _tray = null;
        }
    }

    public void RequestShutdown() => _shutdownCts.Cancel();

    private async Task RunPipeServerAsync(CancellationToken cancellationToken)
    {
        // Use a semaphore to limit concurrent connections (prevents resource exhaustion)
        using var connectionLimit = new SemaphoreSlim(10, 10);

        while (!cancellationToken.IsCancellationRequested)
        {
            NamedPipeServerStream? server = null;
            try
            {
                server = DaemonSecurity.CreateSecureServer();
                await server.WaitForConnectionAsync(cancellationToken);
                _lastActivityTime = DateTime.UtcNow;

                // Capture server for the task
                var clientServer = server;
                server = null; // Prevent disposal in finally - task owns it now

                // Handle client asynchronously - allows accepting next connection immediately
                _ = Task.Run(async () =>
                {
                    await connectionLimit.WaitAsync(cancellationToken);
                    try
                    {
                        await HandleClientAsync(clientServer, cancellationToken);
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
            catch
            {
                // Log errors but continue serving
            }
            finally
            {
                if (server != null)
                {
                    try { if (server.IsConnected) server.Disconnect(); } catch { }
                    await server.DisposeAsync();
                }
            }
        }
    }

    private async Task HandleClientAsync(NamedPipeServerStream server, CancellationToken cancellationToken)
    {
        using var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true);
        using var writer = new StreamWriter(server, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

        var requestJson = await reader.ReadLineAsync(cancellationToken);
        if (string.IsNullOrEmpty(requestJson)) return;

        var request = DaemonProtocol.Deserialize<DaemonRequest>(requestJson);
        if (request == null)
        {
            await writer.WriteLineAsync(DaemonProtocol.Serialize(new DaemonResponse { Success = false, ErrorMessage = "Invalid request" }));
            return;
        }

        var response = await ProcessRequestAsync(request);

        // Write response without cancellation token - we need to send response even during shutdown
        await writer.WriteLineAsync(DaemonProtocol.Serialize(response));

        // Ensure response is transmitted before closing pipe
        try
        {
            await server.FlushAsync(CancellationToken.None);
            server.WaitForPipeDrain();
        }
        catch (IOException)
        {
            // Client may have disconnected - that's ok
        }
    }

    private async Task<DaemonResponse> ProcessRequestAsync(DaemonRequest request)
    {
        _lastActivityTime = DateTime.UtcNow;

        try
        {
            // Route command
            var parts = request.Command.Split('.', 2);
            var category = parts[0];
            var action = parts.Length > 1 ? parts[1] : "";

            return category switch
            {
                "daemon" => HandleDaemonCommand(action),
                "session" => HandleSessionCommand(action, request),
                "sheet" => await HandleSheetCommandAsync(action, request),
                "range" => await HandleRangeCommandAsync(action, request),
                "table" => await HandleTableCommandAsync(action, request),
                "powerquery" => await HandlePowerQueryCommandAsync(action, request),
                "pivottable" => await HandlePivotTableCommandAsync(action, request),
                "chart" => await HandleChartCommandAsync(action, request),
                "chartconfig" => await HandleChartConfigCommandAsync(action, request),
                "connection" => await HandleConnectionCommandAsync(action, request),
                "namedrange" => await HandleNamedRangeCommandAsync(action, request),
                "conditionalformat" => await HandleConditionalFormatCommandAsync(action, request),
                "vba" => await HandleVbaCommandAsync(action, request),
                "datamodel" => await HandleDataModelCommandAsync(action, request),
                "datamodelrel" => await HandleDataModelRelCommandAsync(action, request),
                "slicer" => await HandleSlicerCommandAsync(action, request),
                _ => new DaemonResponse { Success = false, ErrorMessage = $"Unknown command category: {category}" }
            };
        }
        catch (Exception ex)
        {
            return new DaemonResponse { Success = false, ErrorMessage = ex.Message };
        }
    }

    // === DAEMON COMMANDS ===

    private DaemonResponse HandleDaemonCommand(string action)
    {
        return action switch
        {
            "ping" => new DaemonResponse { Success = true },
            "shutdown" => HandleShutdown(),
            "status" => HandleStatus(),
            _ => new DaemonResponse { Success = false, ErrorMessage = $"Unknown daemon action: {action}" }
        };
    }

    private DaemonResponse HandleShutdown()
    {
        _shutdownCts.Cancel();
        return new DaemonResponse { Success = true };
    }

    private DaemonResponse HandleStatus()
    {
        var status = new DaemonStatus
        {
            Running = true,
            ProcessId = Environment.ProcessId,
            SessionCount = _sessionManager.GetActiveSessions().Count,
            StartTime = _startTime
        };
        return new DaemonResponse { Success = true, Result = JsonSerializer.Serialize(status, DaemonProtocol.JsonOptions) };
    }

    // === SESSION COMMANDS ===

    private DaemonResponse HandleSessionCommand(string action, DaemonRequest request)
    {
        return action switch
        {
            "create" => HandleSessionCreate(request),
            "open" => HandleSessionOpen(request),
            "close" => HandleSessionClose(request),
            "save" => HandleSessionSave(request),
            "list" => HandleSessionList(),
            _ => new DaemonResponse { Success = false, ErrorMessage = $"Unknown session action: {action}" }
        };
    }

    private DaemonResponse HandleSessionCreate(DaemonRequest request)
    {
        var args = DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new DaemonResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        var fullPath = Path.GetFullPath(args.FilePath);

        if (File.Exists(fullPath))
        {
            return new DaemonResponse
            {
                Success = false,
                ErrorMessage = $"File already exists: {fullPath}. Use session open to open an existing workbook."
            };
        }

        var extension = Path.GetExtension(fullPath);
        if (!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            return new DaemonResponse
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
            var sessionId = _sessionManager.CreateSessionForNewFile(fullPath, operationTimeout: timeout);

            return new DaemonResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { sessionId, filePath = fullPath }, DaemonProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return new DaemonResponse { Success = false, ErrorMessage = ex.Message };
        }
    }

    private DaemonResponse HandleSessionOpen(DaemonRequest request)
    {
        var args = DeserializeArgs<SessionOpenArgs>(request.Args);
        if (string.IsNullOrWhiteSpace(args?.FilePath))
        {
            return new DaemonResponse { Success = false, ErrorMessage = "filePath is required" };
        }

        try
        {
            TimeSpan? timeout = args.TimeoutSeconds.HasValue
                ? TimeSpan.FromSeconds(args.TimeoutSeconds.Value)
                : null;
            var sessionId = _sessionManager.CreateSession(args.FilePath, operationTimeout: timeout);
            return new DaemonResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { sessionId, filePath = args.FilePath }, DaemonProtocol.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return new DaemonResponse { Success = false, ErrorMessage = ex.Message };
        }
    }

    private DaemonResponse HandleSessionClose(DaemonRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new DaemonResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var args = DeserializeArgs<SessionCloseArgs>(request.Args);
        var closed = _sessionManager.CloseSession(request.SessionId, save: args?.Save ?? false);

        return closed
            ? new DaemonResponse { Success = true }
            : new DaemonResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
    }

    private DaemonResponse HandleSessionSave(DaemonRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SessionId))
        {
            return new DaemonResponse { Success = false, ErrorMessage = "sessionId is required" };
        }

        var batch = _sessionManager.GetSession(request.SessionId);
        if (batch == null)
        {
            return new DaemonResponse { Success = false, ErrorMessage = $"Session '{request.SessionId}' not found" };
        }

        // Check if Excel process is still alive before attempting save
        if (!batch.IsExcelProcessAlive())
        {
            _sessionManager.CloseSession(request.SessionId, save: false, force: true);
            return new DaemonResponse
            {
                Success = false,
                ErrorMessage = $"Excel process for session '{request.SessionId}' has died. Session has been closed. Please create a new session."
            };
        }

        batch.Save();
        return new DaemonResponse { Success = true };
    }

    private DaemonResponse HandleSessionList()
    {
        var sessions = _sessionManager.GetActiveSessions()
            .Select(s => new { sessionId = s.SessionId, filePath = s.FilePath })
            .ToList();

        return new DaemonResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new { sessions }, DaemonProtocol.JsonOptions)
        };
    }

    // === SHEET COMMANDS ===

    private Task<DaemonResponse> HandleSheetCommandAsync(string action, DaemonRequest request)
    {
        if (TryParseAction<WorksheetAction>(action, out var sheetAction))
        {
            if (sheetAction is WorksheetAction.CopyToFile or WorksheetAction.MoveToFile)
            {
                return Task.FromResult(HandleSheetCrossFileCommand(sheetAction, request));
            }

            return WithSessionAsync(request.SessionId, batch => sheetAction switch
            {
                // Lifecycle operations
                WorksheetAction.List => SerializeResult(_sheetCommands.List(batch)),
                WorksheetAction.Create => ExecuteVoid(() => _sheetCommands.Create(batch, GetArg<SheetArgs>(request.Args).SheetName!)),
                WorksheetAction.Rename => ExecuteVoid(() =>
                {
                    var args = GetArg<SheetRenameArgs>(request.Args);
                    _sheetCommands.Rename(batch, args.SheetName!, args.NewName!);
                }),
                WorksheetAction.Delete => ExecuteVoid(() => _sheetCommands.Delete(batch, GetArg<SheetArgs>(request.Args).SheetName!)),
                WorksheetAction.Copy => ExecuteVoid(() =>
                {
                    var args = GetArg<SheetCopyArgs>(request.Args);
                    _sheetCommands.Copy(batch, args.SourceSheet!, args.TargetSheet!);
                }),
                WorksheetAction.Move => ExecuteVoid(() =>
                {
                    var args = GetArg<SheetMoveArgs>(request.Args);
                    _sheetCommands.Move(batch, args.SheetName!, args.BeforeSheet, args.AfterSheet);
                }),

                _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported sheet action: {action}" }
            });
        }

        if (TryParseAction<WorksheetStyleAction>(action, out var styleAction))
        {
            return WithSessionAsync(request.SessionId, batch => styleAction switch
            {
                WorksheetStyleAction.SetTabColor => ExecuteVoid(() =>
                {
                    var args = GetArg<SheetTabColorArgs>(request.Args);
                    _sheetCommands.SetTabColor(batch, args.SheetName!, args.Red ?? 0, args.Green ?? 0, args.Blue ?? 0);
                }),
                WorksheetStyleAction.GetTabColor => SerializeResult(_sheetCommands.GetTabColor(batch, GetArg<SheetArgs>(request.Args).SheetName!)),
                WorksheetStyleAction.ClearTabColor => ExecuteVoid(() => _sheetCommands.ClearTabColor(batch, GetArg<SheetArgs>(request.Args).SheetName!)),

                WorksheetStyleAction.SetVisibility => ExecuteVoid(() =>
                {
                    var args = GetArg<SheetVisibilityArgs>(request.Args);
                    var visibility = args.Visibility?.ToLowerInvariant() switch
                    {
                        "visible" => SheetVisibility.Visible,
                        "hidden" => SheetVisibility.Hidden,
                        "veryhidden" or "very-hidden" => SheetVisibility.VeryHidden,
                        _ => SheetVisibility.Visible
                    };
                    _sheetCommands.SetVisibility(batch, args.SheetName!, visibility);
                }),
                WorksheetStyleAction.GetVisibility => SerializeResult(_sheetCommands.GetVisibility(batch, GetArg<SheetArgs>(request.Args).SheetName!)),
                WorksheetStyleAction.Show => ExecuteVoid(() => _sheetCommands.Show(batch, GetArg<SheetArgs>(request.Args).SheetName!)),
                WorksheetStyleAction.Hide => ExecuteVoid(() => _sheetCommands.Hide(batch, GetArg<SheetArgs>(request.Args).SheetName!)),
                WorksheetStyleAction.VeryHide => ExecuteVoid(() => _sheetCommands.VeryHide(batch, GetArg<SheetArgs>(request.Args).SheetName!)),

                _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported sheet style action: {action}" }
            });
        }

        return Task.FromResult(new DaemonResponse { Success = false, ErrorMessage = $"Unknown sheet action: {action}" });
    }

    private DaemonResponse HandleSheetCrossFileCommand(WorksheetAction action, DaemonRequest request)
    {
        try
        {
            return action switch
            {
                WorksheetAction.CopyToFile => ExecuteVoid(() =>
                {
                    var args = GetArg<SheetCopyToFileArgs>(request.Args);
                    _sheetCommands.CopyToFile(args.SourceFile!, args.SourceSheet!, args.TargetFile!, args.TargetSheetName, args.BeforeSheet, args.AfterSheet);
                }),
                WorksheetAction.MoveToFile => ExecuteVoid(() =>
                {
                    var args = GetArg<SheetMoveToFileArgs>(request.Args);
                    _sheetCommands.MoveToFile(args.SourceFile!, args.SourceSheet!, args.TargetFile!, args.BeforeSheet, args.AfterSheet);
                }),
                _ => new DaemonResponse { Success = false, ErrorMessage = $"Unknown cross-file sheet action: {action}" }
            };
        }
        catch (Exception ex)
        {
            return new DaemonResponse { Success = false, ErrorMessage = ex.Message };
        }
    }

    // === RANGE COMMANDS (stub - implement as needed) ===

    private Task<DaemonResponse> HandleRangeCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<RangeAction>(action, out var rangeAction))
            {
                return rangeAction switch
                {
                    RangeAction.GetValues => ExecuteRangeGetValues(batch, request),
                    RangeAction.SetValues => ExecuteRangeSetValues(batch, request),
                    RangeAction.GetUsedRange => ExecuteRangeGetUsedRange(batch, request),
                    RangeAction.GetCurrentRegion => ExecuteRangeGetCurrentRegion(batch, request),
                    RangeAction.GetInfo => ExecuteRangeGetInfo(batch, request),
                    RangeAction.GetFormulas => ExecuteRangeGetFormulas(batch, request),
                    RangeAction.SetFormulas => ExecuteRangeSetFormulas(batch, request),
                    RangeAction.ClearAll => ExecuteRangeClearAll(batch, request),
                    RangeAction.ClearContents => ExecuteRangeClearContents(batch, request),
                    RangeAction.ClearFormats => ExecuteRangeClearFormats(batch, request),
                    RangeAction.Copy => ExecuteRangeCopy(batch, request),
                    RangeAction.CopyValues => ExecuteRangeCopyValues(batch, request),
                    RangeAction.CopyFormulas => ExecuteRangeCopyFormulas(batch, request),
                    RangeAction.GetNumberFormats => ExecuteRangeGetNumberFormats(batch, request),
                    RangeAction.SetNumberFormat => ExecuteRangeSetNumberFormat(batch, request),
                    RangeAction.SetNumberFormats => ExecuteRangeSetNumberFormats(batch, request),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported range action: {action}" }
                };
            }

            if (TryParseAction<RangeEditAction>(action, out var editAction))
            {
                return editAction switch
                {
                    RangeEditAction.InsertCells => ExecuteRangeInsertCells(batch, request),
                    RangeEditAction.DeleteCells => ExecuteRangeDeleteCells(batch, request),
                    RangeEditAction.InsertRows => ExecuteRangeInsertRows(batch, request),
                    RangeEditAction.DeleteRows => ExecuteRangeDeleteRows(batch, request),
                    RangeEditAction.InsertColumns => ExecuteRangeInsertColumns(batch, request),
                    RangeEditAction.DeleteColumns => ExecuteRangeDeleteColumns(batch, request),
                    RangeEditAction.Find => ExecuteRangeFind(batch, request),
                    RangeEditAction.Replace => ExecuteRangeReplace(batch, request),
                    RangeEditAction.Sort => ExecuteRangeSort(batch, request),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported range edit action: {action}" }
                };
            }

            if (TryParseAction<RangeFormatAction>(action, out var formatAction))
            {
                return formatAction switch
                {
                    RangeFormatAction.SetStyle => ExecuteRangeSetStyle(batch, request),
                    RangeFormatAction.GetStyle => ExecuteRangeGetStyle(batch, request),
                    RangeFormatAction.FormatRange => ExecuteRangeFormatRange(batch, request),
                    RangeFormatAction.ValidateRange => ExecuteRangeValidateRange(batch, request),
                    RangeFormatAction.GetValidation => ExecuteRangeGetValidation(batch, request),
                    RangeFormatAction.RemoveValidation => ExecuteRangeRemoveValidation(batch, request),
                    RangeFormatAction.AutoFitColumns => ExecuteRangeAutoFitColumns(batch, request),
                    RangeFormatAction.AutoFitRows => ExecuteRangeAutoFitRows(batch, request),
                    RangeFormatAction.MergeCells => ExecuteRangeMergeCells(batch, request),
                    RangeFormatAction.UnmergeCells => ExecuteRangeUnmergeCells(batch, request),
                    RangeFormatAction.GetMergeInfo => ExecuteRangeGetMergeInfo(batch, request),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported range format action: {action}" }
                };
            }

            if (TryParseAction<RangeLinkAction>(action, out var linkAction))
            {
                return linkAction switch
                {
                    RangeLinkAction.AddHyperlink => ExecuteRangeAddHyperlink(batch, request),
                    RangeLinkAction.RemoveHyperlink => ExecuteRangeRemoveHyperlink(batch, request),
                    RangeLinkAction.ListHyperlinks => ExecuteRangeListHyperlinks(batch, request),
                    RangeLinkAction.GetHyperlink => ExecuteRangeGetHyperlink(batch, request),
                    RangeLinkAction.SetCellLock => ExecuteRangeSetCellLock(batch, request),
                    RangeLinkAction.GetCellLock => ExecuteRangeGetCellLock(batch, request),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported range link action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown range action: {action}" };
        });
    }

    private DaemonResponse ExecuteRangeGetValues(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.GetValues(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeSetValues(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeSetValuesArgs>(request.Args);
        var result = _rangeCommands.SetValues(batch, args.SheetName!, args.Range!, args.Values!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeGetUsedRange(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<SheetArgs>(request.Args);
        var result = _rangeCommands.GetUsedRange(batch, args.SheetName!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeGetCurrentRegion(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeCellArgs>(request.Args);
        var result = _rangeCommands.GetCurrentRegion(batch, args.SheetName!, args.CellAddress!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeGetInfo(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.GetInfo(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeGetFormulas(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.GetFormulas(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeSetFormulas(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeSetFormulasArgs>(request.Args);
        var result = _rangeCommands.SetFormulas(batch, args.SheetName!, args.Range!, args.Formulas!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeClearAll(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.ClearAll(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeClearContents(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.ClearContents(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeClearFormats(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.ClearFormats(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeCopy(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeCopyArgs>(request.Args);
        var result = _rangeCommands.Copy(batch, args.SourceSheet!, args.SourceRange!, args.TargetSheet!, args.TargetRange!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeCopyValues(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeCopyArgs>(request.Args);
        var result = _rangeCommands.CopyValues(batch, args.SourceSheet!, args.SourceRange!, args.TargetSheet!, args.TargetRange!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeCopyFormulas(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeCopyArgs>(request.Args);
        var result = _rangeCommands.CopyFormulas(batch, args.SourceSheet!, args.SourceRange!, args.TargetSheet!, args.TargetRange!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeInsertCells(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeInsertCellsArgs>(request.Args);
        var shift = args.ShiftDirection?.ToLowerInvariant() == "right"
            ? Core.Commands.Range.InsertShiftDirection.Right
            : Core.Commands.Range.InsertShiftDirection.Down;
        var result = _rangeCommands.InsertCells(batch, args.SheetName!, args.Range!, shift);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeDeleteCells(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeDeleteCellsArgs>(request.Args);
        var shift = args.ShiftDirection?.ToLowerInvariant() == "left"
            ? Core.Commands.Range.DeleteShiftDirection.Left
            : Core.Commands.Range.DeleteShiftDirection.Up;
        var result = _rangeCommands.DeleteCells(batch, args.SheetName!, args.Range!, shift);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeInsertRows(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.InsertRows(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeDeleteRows(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.DeleteRows(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeInsertColumns(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.InsertColumns(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeDeleteColumns(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.DeleteColumns(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeFind(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeFindArgs>(request.Args);
        var options = new Core.Commands.Range.FindOptions
        {
            MatchCase = args.MatchCase ?? false,
            MatchEntireCell = args.MatchEntireCell ?? false,
            SearchFormulas = args.SearchFormulas ?? true,
            SearchValues = args.SearchValues ?? true,
            SearchComments = args.SearchComments ?? false
        };
        var result = _rangeCommands.Find(batch, args.SheetName!, args.Range!, args.SearchValue!, options);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeReplace(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeReplaceArgs>(request.Args);
        var options = new Core.Commands.Range.ReplaceOptions
        {
            MatchCase = args.MatchCase ?? false,
            MatchEntireCell = args.MatchEntireCell ?? false,
            ReplaceAll = args.ReplaceAll ?? true
        };
        _rangeCommands.Replace(batch, args.SheetName!, args.Range!, args.FindValue!, args.ReplaceValue!, options);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeSort(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeSortArgs>(request.Args);
        var sortColumns = args.SortColumns?
            .Select(sc => new Core.Commands.Range.SortColumn { ColumnIndex = sc.ColumnIndex, Ascending = sc.Ascending })
            .ToList() ?? [];
        _rangeCommands.Sort(batch, args.SheetName!, args.Range!, sortColumns, args.HasHeaders ?? true);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeAddHyperlink(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeHyperlinkArgs>(request.Args);
        var result = _rangeCommands.AddHyperlink(batch, args.SheetName!, args.CellAddress!, args.Url!, args.DisplayText, args.Tooltip);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeRemoveHyperlink(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.RemoveHyperlink(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeListHyperlinks(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<SheetArgs>(request.Args);
        var result = _rangeCommands.ListHyperlinks(batch, args.SheetName!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeGetHyperlink(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeCellArgs>(request.Args);
        var result = _rangeCommands.GetHyperlink(batch, args.SheetName!, args.CellAddress!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeGetNumberFormats(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.GetNumberFormats(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeSetNumberFormat(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeNumberFormatArgs>(request.Args);
        var result = _rangeCommands.SetNumberFormat(batch, args.SheetName!, args.Range!, args.FormatCode!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeSetNumberFormats(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeNumberFormatsArgs>(request.Args);
        var result = _rangeCommands.SetNumberFormats(batch, args.SheetName!, args.Range!, args.Formats!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeSetStyle(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeStyleArgs>(request.Args);
        _rangeCommands.SetStyle(batch, args.SheetName!, args.Range!, args.StyleName!);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeGetStyle(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.GetStyle(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeFormatRange(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeFormatArgs>(request.Args);
        _rangeCommands.FormatRange(
            batch, args.SheetName!, args.Range!,
            args.FontName, args.FontSize, args.Bold, args.Italic, args.Underline,
            args.FontColor, args.FillColor, args.BorderStyle, args.BorderColor, args.BorderWeight,
            args.HorizontalAlignment, args.VerticalAlignment, args.WrapText, args.Orientation);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeValidateRange(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeValidationArgs>(request.Args);
        _rangeCommands.ValidateRange(
            batch, args.SheetName!, args.Range!,
            args.ValidationType!, args.ValidationOperator, args.Formula1, args.Formula2,
            args.ShowInputMessage, args.InputTitle, args.InputMessage,
            args.ShowErrorAlert, args.ErrorStyle, args.ErrorTitle, args.ErrorMessage,
            args.IgnoreBlank, args.ShowDropdown);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeGetValidation(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.GetValidation(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeRemoveValidation(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        _rangeCommands.RemoveValidation(batch, args.SheetName!, args.Range!);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeAutoFitColumns(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        _rangeCommands.AutoFitColumns(batch, args.SheetName!, args.Range!);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeAutoFitRows(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        _rangeCommands.AutoFitRows(batch, args.SheetName!, args.Range!);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeMergeCells(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        _rangeCommands.MergeCells(batch, args.SheetName!, args.Range!);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeUnmergeCells(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        _rangeCommands.UnmergeCells(batch, args.SheetName!, args.Range!);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeGetMergeInfo(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.GetMergeInfo(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    private DaemonResponse ExecuteRangeSetCellLock(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeLockArgs>(request.Args);
        _rangeCommands.SetCellLock(batch, args.SheetName!, args.Range!, args.Locked ?? true);
        return ExecuteVoid(() => { });
    }

    private DaemonResponse ExecuteRangeGetCellLock(IExcelBatch batch, DaemonRequest request)
    {
        var args = GetArg<RangeArgs>(request.Args);
        var result = _rangeCommands.GetCellLock(batch, args.SheetName!, args.Range!);
        return SerializeResult(result);
    }

    // === OTHER COMMAND CATEGORIES ===

    private Task<DaemonResponse> HandleTableCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<TableAction>(action, out var tableAction))
            {
                return tableAction switch
                {
                    TableAction.List => SerializeResult(_tableCommands.List(batch)),
                    TableAction.Create => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableCreateArgs>(request.Args);
                        _tableCommands.Create(batch, a.SheetName!, a.TableName!, a.Range!, a.HasHeaders ?? true, a.TableStyle);
                    }),
                    TableAction.Read => SerializeResult(_tableCommands.Read(batch, GetArg<TableArgs>(request.Args).TableName!)),
                    TableAction.Rename => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableRenameArgs>(request.Args);
                        _tableCommands.Rename(batch, a.TableName!, a.NewName!);
                    }),
                    TableAction.Delete => ExecuteVoid(() => _tableCommands.Delete(batch, GetArg<TableArgs>(request.Args).TableName!)),
                    TableAction.Resize => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableResizeArgs>(request.Args);
                        _tableCommands.Resize(batch, a.TableName!, a.NewRange!);
                    }),
                    TableAction.SetStyle => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableStyleArgs>(request.Args);
                        _tableCommands.SetStyle(batch, a.TableName!, a.TableStyle!);
                    }),
                    TableAction.ToggleTotals => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableToggleTotalsArgs>(request.Args);
                        _tableCommands.ToggleTotals(batch, a.TableName!, a.ShowTotals ?? false);
                    }),
                    TableAction.SetColumnTotal => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableColumnTotalArgs>(request.Args);
                        _tableCommands.SetColumnTotal(batch, a.TableName!, a.ColumnName!, a.TotalFunction!);
                    }),
                    TableAction.Append => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableAppendArgs>(request.Args);
                        _tableCommands.Append(batch, a.TableName!, a.Rows!);
                    }),
                    TableAction.GetData => SerializeResult(_tableCommands.GetData(batch, GetArg<TableDataArgs>(request.Args).TableName!, GetArg<TableDataArgs>(request.Args).VisibleOnly ?? false)),
                    TableAction.AddToDataModel => ExecuteVoid(() => _tableCommands.AddToDataModel(batch, GetArg<TableArgs>(request.Args).TableName!)),
                    TableAction.CreateFromDax => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableDaxArgs>(request.Args);
                        _tableCommands.CreateFromDax(batch, a.SheetName!, a.TableName!, a.DaxQuery!, a.TargetCell);
                    }),
                    TableAction.UpdateDax => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableUpdateDaxArgs>(request.Args);
                        _tableCommands.UpdateDax(batch, a.TableName!, a.DaxQuery!);
                    }),
                    TableAction.GetDax => SerializeResult(_tableCommands.GetDax(batch, GetArg<TableArgs>(request.Args).TableName!)),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported table action: {action}" }
                };
            }

            if (TryParseAction<TableColumnAction>(action, out var columnAction))
            {
                return columnAction switch
                {
                    TableColumnAction.ApplyFilter or TableColumnAction.ApplyFilterValues => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableFilterArgs>(request.Args);
                        if (a.Criteria != null)
                            _tableCommands.ApplyFilter(batch, a.TableName!, a.ColumnName!, a.Criteria);
                        else if (a.Values != null)
                            _tableCommands.ApplyFilter(batch, a.TableName!, a.ColumnName!, a.Values);
                    }),
                    TableColumnAction.ClearFilters => ExecuteVoid(() => _tableCommands.ClearFilters(batch, GetArg<TableArgs>(request.Args).TableName!)),
                    TableColumnAction.GetFilters => SerializeResult(_tableCommands.GetFilters(batch, GetArg<TableArgs>(request.Args).TableName!)),
                    TableColumnAction.AddColumn => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableAddColumnArgs>(request.Args);
                        _tableCommands.AddColumn(batch, a.TableName!, a.ColumnName!, a.Position);
                    }),
                    TableColumnAction.RemoveColumn => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableColumnArgs>(request.Args);
                        _tableCommands.RemoveColumn(batch, a.TableName!, a.ColumnName!);
                    }),
                    TableColumnAction.RenameColumn => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableRenameColumnArgs>(request.Args);
                        _tableCommands.RenameColumn(batch, a.TableName!, a.OldName!, a.NewName!);
                    }),
                    TableColumnAction.Sort => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableSortArgs>(request.Args);
                        _tableCommands.Sort(batch, a.TableName!, a.ColumnName!, a.Ascending ?? true);
                    }),
                    TableColumnAction.SortMulti => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableSortMultiArgs>(request.Args);
                        var sortColumns = JsonSerializer.Deserialize<List<TableSortColumn>>(a.SortColumnsJson!, DaemonProtocol.JsonOptions)
                            ?? throw new ArgumentException("sortColumnsJson must be a non-empty array");
                        _tableCommands.Sort(batch, a.TableName!, sortColumns);
                    }),
                    TableColumnAction.GetStructuredReference => SerializeResult(() =>
                    {
                        var a = GetArg<TableStructuredRefArgs>(request.Args);
                        var region = ParseTableRegion(a.Region);
                        return _tableCommands.GetStructuredReference(batch, a.TableName!, region, a.ColumnName);
                    }),
                    TableColumnAction.GetColumnNumberFormat => SerializeResult(_tableCommands.GetColumnNumberFormat(batch, GetArg<TableColumnArgs>(request.Args).TableName!, GetArg<TableColumnArgs>(request.Args).ColumnName!)),
                    TableColumnAction.SetColumnNumberFormat => ExecuteVoid(() =>
                    {
                        var a = GetArg<TableColumnFormatArgs>(request.Args);
                        _tableCommands.SetColumnNumberFormat(batch, a.TableName!, a.ColumnName!, a.FormatCode!);
                    }),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported table column action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown table action: {action}" };
        });
    }

    private Task<DaemonResponse> HandlePowerQueryCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<PowerQueryAction>(action, out var pqAction))
            {
                return pqAction switch
                {
                    PowerQueryAction.List => SerializeResult(_powerQueryCommands.List(batch)),
                    PowerQueryAction.View => SerializeResult(_powerQueryCommands.View(batch, GetArg<PowerQueryArgs>(request.Args).QueryName!)),
                    PowerQueryAction.Create => ExecuteVoid(() =>
                    {
                        var a = GetArg<PowerQueryCreateArgs>(request.Args);
                        var loadMode = ParseLoadMode(a.LoadDestination);
                        _powerQueryCommands.Create(batch, a.QueryName!, a.MCode!, loadMode, a.TargetSheet, a.TargetCellAddress);
                    }),
                    PowerQueryAction.Update => ExecuteVoid(() =>
                    {
                        var a = GetArg<PowerQueryUpdateArgs>(request.Args);
                        _powerQueryCommands.Update(batch, a.QueryName!, a.MCode!, a.Refresh ?? true);
                    }),
                    PowerQueryAction.Rename => SerializeResult(_powerQueryCommands.Rename(batch, GetArg<PowerQueryRenameArgs>(request.Args).OldName!, GetArg<PowerQueryRenameArgs>(request.Args).NewName!)),
                    PowerQueryAction.Delete => ExecuteVoid(() => _powerQueryCommands.Delete(batch, GetArg<PowerQueryArgs>(request.Args).QueryName!)),
                    PowerQueryAction.Refresh => SerializeResult(_powerQueryCommands.Refresh(batch, GetArg<PowerQueryArgs>(request.Args).QueryName!, TimeSpan.FromMinutes(5))),
                    PowerQueryAction.RefreshAll => ExecuteVoid(() => _powerQueryCommands.RefreshAll(batch)),
                    PowerQueryAction.LoadTo => ExecuteVoid(() =>
                    {
                        var a = GetArg<PowerQueryLoadToArgs>(request.Args);
                        var loadMode = ParseLoadMode(a.LoadDestination);
                        _powerQueryCommands.LoadTo(batch, a.QueryName!, loadMode, a.TargetSheet, a.TargetCellAddress);
                    }),
                    PowerQueryAction.GetLoadConfig => SerializeResult(_powerQueryCommands.GetLoadConfig(batch, GetArg<PowerQueryArgs>(request.Args).QueryName!)),
                    PowerQueryAction.Unload => SerializeResult(_powerQueryCommands.Unload(batch, GetArg<PowerQueryArgs>(request.Args).QueryName!)),
                    PowerQueryAction.Evaluate => SerializeResult(_powerQueryCommands.Evaluate(batch, GetArg<PowerQueryEvaluateArgs>(request.Args).MCode!)),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported powerquery action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown powerquery action: {action}" };
        });
    }

    private static PowerQueryLoadMode ParseLoadMode(string? loadDestination)
    {
        return loadDestination?.ToLowerInvariant() switch
        {
            "worksheet" or "table" => PowerQueryLoadMode.LoadToTable,
            "data-model" or "datamodel" => PowerQueryLoadMode.LoadToDataModel,
            "both" => PowerQueryLoadMode.LoadToBoth,
            "connection-only" or "connectiononly" => PowerQueryLoadMode.ConnectionOnly,
            _ => PowerQueryLoadMode.LoadToTable
        };
    }

    private Task<DaemonResponse> HandlePivotTableCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<PivotTableAction>(action, out var pivotAction))
            {
                return pivotAction switch
                {
                    PivotTableAction.List => SerializeResult(_pivotTableCommands.List(batch)),
                    PivotTableAction.Read => SerializeResult(_pivotTableCommands.Read(batch, GetArg<PivotTableArgs>(request.Args).PivotTableName!)),
                    PivotTableAction.CreateFromRange => SerializeResult(() =>
                    {
                        var a = GetArg<PivotTableFromRangeArgs>(request.Args);
                        return _pivotTableCommands.CreateFromRange(batch, a.SourceSheet!, a.SourceRange!, a.DestinationSheet!, a.DestinationCell!, a.PivotTableName!);
                    }),
                    PivotTableAction.CreateFromTable => SerializeResult(() =>
                    {
                        var a = GetArg<PivotTableFromTableArgs>(request.Args);
                        return _pivotTableCommands.CreateFromTable(batch, a.TableName!, a.DestinationSheet!, a.DestinationCell!, a.PivotTableName!);
                    }),
                    PivotTableAction.CreateFromDataModel => SerializeResult(() =>
                    {
                        var a = GetArg<PivotTableFromDataModelArgs>(request.Args);
                        return _pivotTableCommands.CreateFromDataModel(batch, a.TableName!, a.DestinationSheet!, a.DestinationCell!, a.PivotTableName!);
                    }),
                    PivotTableAction.Delete => SerializeResult(_pivotTableCommands.Delete(batch, GetArg<PivotTableArgs>(request.Args).PivotTableName!)),
                    PivotTableAction.Refresh => SerializeResult(_pivotTableCommands.Refresh(batch, GetArg<PivotTableArgs>(request.Args).PivotTableName!, GetArg<PivotTableRefreshArgs>(request.Args).Timeout)),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported pivottable action: {action}" }
                };
            }

            if (TryParseAction<PivotTableFieldAction>(action, out var fieldAction))
            {
                return fieldAction switch
                {
                    PivotTableFieldAction.ListFields => SerializeResult(_pivotTableCommands.ListFields(batch, GetArg<PivotTableArgs>(request.Args).PivotTableName!)),
                    PivotTableFieldAction.AddRowField => SerializeResult(() =>
                    {
                        var a = GetArg<PivotFieldArgs>(request.Args);
                        return _pivotTableCommands.AddRowField(batch, a.PivotTableName!, a.FieldName!, a.Position);
                    }),
                    PivotTableFieldAction.AddColumnField => SerializeResult(() =>
                    {
                        var a = GetArg<PivotFieldArgs>(request.Args);
                        return _pivotTableCommands.AddColumnField(batch, a.PivotTableName!, a.FieldName!, a.Position);
                    }),
                    PivotTableFieldAction.AddValueField => SerializeResult(() =>
                    {
                        var a = GetArg<PivotValueFieldArgs>(request.Args);
                        var func = ParseAggregationFunction(a.AggregationFunction);
                        return _pivotTableCommands.AddValueField(batch, a.PivotTableName!, a.FieldName!, func, a.CustomName);
                    }),
                    PivotTableFieldAction.AddFilterField => SerializeResult(_pivotTableCommands.AddFilterField(batch, GetArg<PivotFieldArgs>(request.Args).PivotTableName!, GetArg<PivotFieldArgs>(request.Args).FieldName!)),
                    PivotTableFieldAction.RemoveField => SerializeResult(_pivotTableCommands.RemoveField(batch, GetArg<PivotFieldArgs>(request.Args).PivotTableName!, GetArg<PivotFieldArgs>(request.Args).FieldName!)),
                    PivotTableFieldAction.SetFieldFunction => SerializeResult(() =>
                    {
                        var a = GetArg<PivotFieldFunctionArgs>(request.Args);
                        var func = ParseAggregationFunction(a.AggregationFunction);
                        return _pivotTableCommands.SetFieldFunction(batch, a.PivotTableName!, a.FieldName!, func);
                    }),
                    PivotTableFieldAction.SetFieldName => SerializeResult(() =>
                    {
                        var a = GetArg<PivotFieldNameArgs>(request.Args);
                        return _pivotTableCommands.SetFieldName(batch, a.PivotTableName!, a.FieldName!, a.CustomName!);
                    }),
                    PivotTableFieldAction.SetFieldFormat => SerializeResult(() =>
                    {
                        var a = GetArg<PivotFieldFormatArgs>(request.Args);
                        return _pivotTableCommands.SetFieldFormat(batch, a.PivotTableName!, a.FieldName!, a.NumberFormat!);
                    }),
                    PivotTableFieldAction.SetFieldFilter => SerializeResult(() =>
                    {
                        var a = GetArg<PivotFieldFilterArgs>(request.Args);
                        return _pivotTableCommands.SetFieldFilter(batch, a.PivotTableName!, a.FieldName!, a.SelectedValues!);
                    }),
                    PivotTableFieldAction.SortField => SerializeResult(() =>
                    {
                        var a = GetArg<PivotFieldSortArgs>(request.Args);
                        var dir = a.Ascending ?? true ? SortDirection.Ascending : SortDirection.Descending;
                        return _pivotTableCommands.SortField(batch, a.PivotTableName!, a.FieldName!, dir);
                    }),
                    PivotTableFieldAction.GroupByDate => SerializeResult(() =>
                    {
                        var a = GetArg<PivotGroupByDateArgs>(request.Args);
                        var interval = ParseDateGroupingInterval(a.Interval);
                        return _pivotTableCommands.GroupByDate(batch, a.PivotTableName!, a.FieldName!, interval);
                    }),
                    PivotTableFieldAction.GroupByNumeric => SerializeResult(() =>
                    {
                        var a = GetArg<PivotGroupByNumericArgs>(request.Args);
                        return _pivotTableCommands.GroupByNumeric(batch, a.PivotTableName!, a.FieldName!, a.Start, a.End, a.IntervalSize ?? 10);
                    }),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported pivottable field action: {action}" }
                };
            }

            if (TryParseAction<PivotTableCalcAction>(action, out var calcAction))
            {
                return calcAction switch
                {
                    PivotTableCalcAction.ListCalculatedFields => SerializeResult(_pivotTableCommands.ListCalculatedFields(batch, GetArg<PivotTableArgs>(request.Args).PivotTableName!)),
                    PivotTableCalcAction.CreateCalculatedField => SerializeResult(() =>
                    {
                        var a = GetArg<PivotCalculatedFieldArgs>(request.Args);
                        return _pivotTableCommands.CreateCalculatedField(batch, a.PivotTableName!, a.FieldName!, a.Formula!);
                    }),
                    PivotTableCalcAction.DeleteCalculatedField => SerializeResult(_pivotTableCommands.DeleteCalculatedField(batch, GetArg<PivotCalculatedFieldArgs>(request.Args).PivotTableName!, GetArg<PivotCalculatedFieldArgs>(request.Args).FieldName!)),
                    PivotTableCalcAction.ListCalculatedMembers => SerializeResult(_pivotTableCommands.ListCalculatedMembers(batch, GetArg<PivotTableArgs>(request.Args).PivotTableName!)),
                    PivotTableCalcAction.CreateCalculatedMember => SerializeResult(() =>
                    {
                        var a = GetArg<PivotCalculatedMemberArgs>(request.Args);
                        var memberType = ParseCalculatedMemberType(a.MemberType);
                        return _pivotTableCommands.CreateCalculatedMember(batch, a.PivotTableName!, a.MemberName!, a.Formula!, memberType, a.SolveOrder ?? 0, a.DisplayFolder, a.NumberFormat);
                    }),
                    PivotTableCalcAction.DeleteCalculatedMember => SerializeResult(_pivotTableCommands.DeleteCalculatedMember(batch, GetArg<PivotCalculatedMemberArgs>(request.Args).PivotTableName!, GetArg<PivotCalculatedMemberArgs>(request.Args).MemberName!)),
                    PivotTableCalcAction.SetLayout => SerializeResult(_pivotTableCommands.SetLayout(batch, GetArg<PivotLayoutArgs>(request.Args).PivotTableName!, GetArg<PivotLayoutArgs>(request.Args).LayoutType ?? 1)),
                    PivotTableCalcAction.SetSubtotals => SerializeResult(() =>
                    {
                        var a = GetArg<PivotSubtotalsArgs>(request.Args);
                        return _pivotTableCommands.SetSubtotals(batch, a.PivotTableName!, a.FieldName!, a.ShowSubtotals ?? true);
                    }),
                    PivotTableCalcAction.SetGrandTotals => SerializeResult(_pivotTableCommands.SetGrandTotals(batch, GetArg<PivotGrandTotalsArgs>(request.Args).PivotTableName!, GetArg<PivotGrandTotalsArgs>(request.Args).ShowRowGrandTotals ?? true, GetArg<PivotGrandTotalsArgs>(request.Args).ShowColumnGrandTotals ?? true)),
                    PivotTableCalcAction.GetData => SerializeResult(_pivotTableCommands.GetData(batch, GetArg<PivotTableArgs>(request.Args).PivotTableName!)),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported pivottable calc action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown pivottable action: {action}" };
        });
    }

    private static AggregationFunction ParseAggregationFunction(string? func)
    {
        return func?.ToLowerInvariant() switch
        {
            "sum" => AggregationFunction.Sum,
            "count" => AggregationFunction.Count,
            "average" or "avg" => AggregationFunction.Average,
            "max" => AggregationFunction.Max,
            "min" => AggregationFunction.Min,
            "product" => AggregationFunction.Product,
            "countnums" or "countnumbers" => AggregationFunction.CountNumbers,
            "stddev" => AggregationFunction.StdDev,
            "stddevp" => AggregationFunction.StdDevP,
            "var" => AggregationFunction.Var,
            "varp" => AggregationFunction.VarP,
            _ => AggregationFunction.Sum
        };
    }

    private static DateGroupingInterval ParseDateGroupingInterval(string? interval)
    {
        return interval?.ToLowerInvariant() switch
        {
            "days" or "day" => DateGroupingInterval.Days,
            "months" or "month" => DateGroupingInterval.Months,
            "quarters" or "quarter" => DateGroupingInterval.Quarters,
            "years" or "year" => DateGroupingInterval.Years,
            _ => DateGroupingInterval.Months
        };
    }

    private static CalculatedMemberType ParseCalculatedMemberType(string? memberType)
    {
        return memberType?.ToLowerInvariant() switch
        {
            "member" => CalculatedMemberType.Member,
            "set" => CalculatedMemberType.Set,
            "measure" => CalculatedMemberType.Measure,
            _ => CalculatedMemberType.Member
        };
    }

    private static TableRegion ParseTableRegion(string? region)
    {
        return region?.ToLowerInvariant() switch
        {
            "all" => TableRegion.All,
            "data" => TableRegion.Data,
            "headers" or "header" => TableRegion.Headers,
            "totals" or "total" => TableRegion.Totals,
            _ => TableRegion.Data
        };
    }

    private Task<DaemonResponse> HandleChartCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<ChartAction>(action, out var chartAction))
            {
                return chartAction switch
                {
                    ChartAction.List => SerializeResult(_chartCommands.List(batch)),
                    ChartAction.Read => SerializeResult(_chartCommands.Read(batch, GetArg<ChartArgs>(request.Args).ChartName!)),
                    ChartAction.CreateFromRange => SerializeResult(() =>
                    {
                        var a = GetArg<ChartFromRangeArgs>(request.Args);
                        var chartType = ParseChartType(a.ChartType);
                        return _chartCommands.CreateFromRange(batch, a.SheetName!, a.SourceRange!, chartType, a.Left ?? 0, a.Top ?? 0, a.Width ?? 400, a.Height ?? 300, a.ChartName);
                    }),
                    ChartAction.CreateFromTable => SerializeResult(() =>
                    {
                        var a = GetArg<ChartFromTableArgs>(request.Args);
                        var chartType = ParseChartType(a.ChartType);
                        return _chartCommands.CreateFromTable(batch, a.TableName!, a.SheetName!, chartType, a.Left ?? 0, a.Top ?? 0, a.Width ?? 400, a.Height ?? 300, a.ChartName);
                    }),
                    ChartAction.CreateFromPivotTable => SerializeResult(() =>
                    {
                        var a = GetArg<ChartFromPivotArgs>(request.Args);
                        var chartType = ParseChartType(a.ChartType);
                        return _chartCommands.CreateFromPivotTable(batch, a.PivotTableName!, a.SheetName!, chartType, a.Left ?? 0, a.Top ?? 0, a.Width ?? 400, a.Height ?? 300, a.ChartName);
                    }),
                    ChartAction.Delete => ExecuteVoid(() => _chartCommands.Delete(batch, GetArg<ChartArgs>(request.Args).ChartName!)),
                    ChartAction.Move => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartMoveArgs>(request.Args);
                        _chartCommands.Move(batch, a.ChartName!, a.Left, a.Top, a.Width, a.Height);
                    }),
                    ChartAction.FitToRange => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartFitArgs>(request.Args);
                        _chartCommands.FitToRange(batch, a.ChartName!, a.SheetName!, a.RangeAddress!);
                    }),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported chart action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown chart action: {action}" };
        });
    }

    private Task<DaemonResponse> HandleChartConfigCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<ChartConfigAction>(action, out var configAction))
            {
                return configAction switch
                {
                    ChartConfigAction.SetSourceRange => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartSourceRangeArgs>(request.Args);
                        _chartCommands.SetSourceRange(batch, a.ChartName!, a.SourceRange!);
                    }),
                    ChartConfigAction.AddSeries => SerializeResult(() =>
                    {
                        var a = GetArg<ChartAddSeriesArgs>(request.Args);
                        return _chartCommands.AddSeries(batch, a.ChartName!, a.SeriesName!, a.ValuesRange!, a.CategoryRange);
                    }),
                    ChartConfigAction.RemoveSeries => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartRemoveSeriesArgs>(request.Args);
                        _chartCommands.RemoveSeries(batch, a.ChartName!, a.SeriesIndex ?? 1);
                    }),
                    ChartConfigAction.SetChartType => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartTypeArgs>(request.Args);
                        _chartCommands.SetChartType(batch, a.ChartName!, ParseChartType(a.ChartType));
                    }),
                    ChartConfigAction.SetTitle => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartTitleArgs>(request.Args);
                        _chartCommands.SetTitle(batch, a.ChartName!, a.Title!);
                    }),
                    ChartConfigAction.SetAxisTitle => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartAxisTitleArgs>(request.Args);
                        _chartCommands.SetAxisTitle(batch, a.ChartName!, ParseAxisType(a.Axis), a.Title!);
                    }),
                    ChartConfigAction.GetAxisNumberFormat => SerializeResult(new { numberFormat = _chartCommands.GetAxisNumberFormat(batch, GetArg<ChartAxisArgs>(request.Args).ChartName!, ParseAxisType(GetArg<ChartAxisArgs>(request.Args).Axis)) }),
                    ChartConfigAction.SetAxisNumberFormat => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartAxisFormatArgs>(request.Args);
                        _chartCommands.SetAxisNumberFormat(batch, a.ChartName!, ParseAxisType(a.Axis), a.NumberFormat!);
                    }),
                    ChartConfigAction.ShowLegend => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartLegendArgs>(request.Args);
                        _chartCommands.ShowLegend(batch, a.ChartName!, a.Visible ?? true, ParseLegendPosition(a.LegendPosition));
                    }),
                    ChartConfigAction.SetStyle => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartStyleArgs>(request.Args);
                        _chartCommands.SetStyle(batch, a.ChartName!, a.StyleId ?? 1);
                    }),
                    ChartConfigAction.SetPlacement => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartPlacementArgs>(request.Args);
                        _chartCommands.SetPlacement(batch, a.ChartName!, a.Placement ?? 2);
                    }),
                    ChartConfigAction.SetDataLabels => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartDataLabelsArgs>(request.Args);
                        _chartCommands.SetDataLabels(batch, a.ChartName!, a.ShowValue, a.ShowPercentage, a.ShowSeriesName, a.ShowCategoryName, null, a.Separator, ParseDataLabelPosition(a.LabelPosition), a.SeriesIndex);
                    }),
                    ChartConfigAction.GetAxisScale => SerializeResult(_chartCommands.GetAxisScale(batch, GetArg<ChartAxisArgs>(request.Args).ChartName!, ParseAxisType(GetArg<ChartAxisArgs>(request.Args).Axis))),
                    ChartConfigAction.SetAxisScale => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartAxisScaleArgs>(request.Args);
                        _chartCommands.SetAxisScale(batch, a.ChartName!, ParseAxisType(a.Axis), a.MinimumScale, a.MaximumScale, a.MajorUnit, a.MinorUnit);
                    }),
                    ChartConfigAction.GetGridlines => SerializeResult(_chartCommands.GetGridlines(batch, GetArg<ChartArgs>(request.Args).ChartName!)),
                    ChartConfigAction.SetGridlines => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartGridlinesArgs>(request.Args);
                        _chartCommands.SetGridlines(batch, a.ChartName!, ParseAxisType(a.Axis), a.ShowMajor, a.ShowMinor);
                    }),
                    ChartConfigAction.SetSeriesFormat => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartSeriesFormatArgs>(request.Args);
                        _chartCommands.SetSeriesFormat(batch, a.ChartName!, a.SeriesIndex ?? 1, ParseMarkerStyle(a.MarkerStyle), a.MarkerSize, a.MarkerBackgroundColor, a.MarkerForegroundColor, null);
                    }),
                    ChartConfigAction.ListTrendlines => SerializeResult(_chartCommands.ListTrendlines(batch, GetArg<ChartSeriesArgs>(request.Args).ChartName!, GetArg<ChartSeriesArgs>(request.Args).SeriesIndex ?? 1)),
                    ChartConfigAction.AddTrendline => SerializeResult(() =>
                    {
                        var a = GetArg<ChartAddTrendlineArgs>(request.Args);
                        return _chartCommands.AddTrendline(batch, a.ChartName!, a.SeriesIndex ?? 1, ParseTrendlineType(a.TrendlineType), null, null, null, null, null, a.DisplayEquation ?? false, a.DisplayRSquared ?? false, a.TrendlineName);
                    }),
                    ChartConfigAction.DeleteTrendline => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartDeleteTrendlineArgs>(request.Args);
                        _chartCommands.DeleteTrendline(batch, a.ChartName!, a.SeriesIndex ?? 1, a.TrendlineIndex ?? 1);
                    }),
                    ChartConfigAction.SetTrendline => ExecuteVoid(() =>
                    {
                        var a = GetArg<ChartSetTrendlineArgs>(request.Args);
                        _chartCommands.SetTrendline(batch, a.ChartName!, a.SeriesIndex ?? 1, a.TrendlineIndex ?? 1, null, null, null, a.DisplayEquation, a.DisplayRSquared, a.TrendlineName);
                    }),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported chartconfig action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown chartconfig action: {action}" };
        });
    }

    private static ChartType ParseChartType(string? chartType)
    {
        return chartType?.ToLowerInvariant() switch
        {
            "column" or "columnclustered" => ChartType.ColumnClustered,
            "columnstacked" => ChartType.ColumnStacked,
            "bar" or "barclustered" => ChartType.BarClustered,
            "barstacked" => ChartType.BarStacked,
            "line" => ChartType.Line,
            "linemarkers" => ChartType.LineMarkers,
            "pie" => ChartType.Pie,
            "doughnut" => ChartType.Doughnut,
            "area" => ChartType.Area,
            "areastacked" => ChartType.AreaStacked,
            "scatter" or "xyscatter" => ChartType.XYScatter,
            "scatterlines" or "xyscatterlines" => ChartType.XYScatterLines,
            _ => ChartType.ColumnClustered
        };
    }

    private static ChartAxisType ParseAxisType(string? axis)
    {
        return axis?.ToLowerInvariant() switch
        {
            "category" or "x" => ChartAxisType.Category,
            "value" or "y" => ChartAxisType.Value,

            _ => ChartAxisType.Value
        };
    }

    private static LegendPosition? ParseLegendPosition(string? position)
    {
        return position?.ToLowerInvariant() switch
        {
            "bottom" => LegendPosition.Bottom,
            "corner" => LegendPosition.Corner,
            "left" => LegendPosition.Left,
            "right" => LegendPosition.Right,
            "top" => LegendPosition.Top,
            _ => null
        };
    }

    private static DataLabelPosition? ParseDataLabelPosition(string? position)
    {
        return position?.ToLowerInvariant() switch
        {
            "center" => DataLabelPosition.Center,
            "insidebase" => DataLabelPosition.InsideBase,
            "insideend" => DataLabelPosition.InsideEnd,
            "outsideend" => DataLabelPosition.OutsideEnd,
            "left" => DataLabelPosition.Left,
            "right" => DataLabelPosition.Right,
            "above" => DataLabelPosition.Above,
            "below" => DataLabelPosition.Below,
            _ => null
        };
    }

    private static MarkerStyle? ParseMarkerStyle(string? style)
    {
        return style?.ToLowerInvariant() switch
        {
            "automatic" => MarkerStyle.Automatic,
            "circle" => MarkerStyle.Circle,
            "dash" => MarkerStyle.Dash,
            "diamond" => MarkerStyle.Diamond,
            "dot" => MarkerStyle.Dot,
            "none" => MarkerStyle.None,
            "picture" => MarkerStyle.Picture,
            "plus" => MarkerStyle.Plus,
            "square" => MarkerStyle.Square,
            "star" => MarkerStyle.Star,
            "triangle" => MarkerStyle.Triangle,
            "x" => MarkerStyle.X,
            _ => null
        };
    }

    private static TrendlineType ParseTrendlineType(string? type)
    {
        return type?.ToLowerInvariant() switch
        {
            "exponential" => TrendlineType.Exponential,
            "linear" => TrendlineType.Linear,
            "logarithmic" => TrendlineType.Logarithmic,
            "movingavg" or "movingaverage" => TrendlineType.MovingAverage,
            "polynomial" => TrendlineType.Polynomial,
            "power" => TrendlineType.Power,
            _ => TrendlineType.Linear
        };
    }

    private Task<DaemonResponse> HandleConnectionCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<ConnectionAction>(action, out var connectionAction))
            {
                return connectionAction switch
                {
                    ConnectionAction.List => SerializeResult(_connectionCommands.List(batch)),
                    ConnectionAction.View => SerializeResult(_connectionCommands.View(batch, GetArg<ConnectionArgs>(request.Args).ConnectionName!)),
                    ConnectionAction.Create => ExecuteVoid(() =>
                    {
                        var a = GetArg<ConnectionCreateArgs>(request.Args);
                        _connectionCommands.Create(batch, a.ConnectionName!, a.ConnectionString!, a.CommandText, a.Description);
                    }),
                    ConnectionAction.Refresh => ExecuteVoid(() =>
                    {
                        var a = GetArg<ConnectionRefreshArgs>(request.Args);
                        if (a.TimeoutSeconds.HasValue)
                            _connectionCommands.Refresh(batch, a.ConnectionName!, TimeSpan.FromSeconds(a.TimeoutSeconds.Value));
                        else
                            _connectionCommands.Refresh(batch, a.ConnectionName!);
                    }),
                    ConnectionAction.Delete => ExecuteVoid(() => _connectionCommands.Delete(batch, GetArg<ConnectionArgs>(request.Args).ConnectionName!)),
                    ConnectionAction.LoadTo => ExecuteVoid(() =>
                    {
                        var a = GetArg<ConnectionLoadToArgs>(request.Args);
                        _connectionCommands.LoadTo(batch, a.ConnectionName!, a.SheetName!);
                    }),
                    ConnectionAction.GetProperties => SerializeResult(_connectionCommands.GetProperties(batch, GetArg<ConnectionArgs>(request.Args).ConnectionName!)),
                    ConnectionAction.SetProperties => ExecuteVoid(() =>
                    {
                        var a = GetArg<ConnectionSetPropertiesArgs>(request.Args);
                        _connectionCommands.SetProperties(batch, a.ConnectionName!, a.ConnectionString, a.CommandText, a.Description, a.BackgroundQuery, a.RefreshOnFileOpen, a.SavePassword, a.RefreshPeriod);
                    }),
                    ConnectionAction.Test => ExecuteVoid(() => _connectionCommands.Test(batch, GetArg<ConnectionArgs>(request.Args).ConnectionName!)),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported connection action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown connection action: {action}" };
        });
    }

    private Task<DaemonResponse> HandleNamedRangeCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<NamedRangeAction>(action, out var namedRangeAction))
            {
                return namedRangeAction switch
                {
                    NamedRangeAction.List => SerializeResult(_namedRangeCommands.List(batch)),
                    NamedRangeAction.Read => SerializeResult(_namedRangeCommands.Read(batch, GetArg<NamedRangeArgs>(request.Args).ParamName!)),
                    NamedRangeAction.Write => ExecuteVoid(() =>
                    {
                        var a = GetArg<NamedRangeWriteArgs>(request.Args);
                        _namedRangeCommands.Write(batch, a.ParamName!, a.Value!);
                    }),
                    NamedRangeAction.Create => ExecuteVoid(() =>
                    {
                        var a = GetArg<NamedRangeCreateArgs>(request.Args);
                        _namedRangeCommands.Create(batch, a.ParamName!, a.Reference!);
                    }),
                    NamedRangeAction.Update => ExecuteVoid(() =>
                    {
                        var a = GetArg<NamedRangeCreateArgs>(request.Args);
                        _namedRangeCommands.Update(batch, a.ParamName!, a.Reference!);
                    }),
                    NamedRangeAction.Delete => ExecuteVoid(() => _namedRangeCommands.Delete(batch, GetArg<NamedRangeArgs>(request.Args).ParamName!)),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported namedrange action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown namedrange action: {action}" };
        });
    }

    private Task<DaemonResponse> HandleConditionalFormatCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<ConditionalFormatAction>(action, out var formatAction))
            {
                return formatAction switch
                {
                    ConditionalFormatAction.AddRule => ExecuteVoid(() =>
                    {
                        var a = GetArg<ConditionalFormatAddArgs>(request.Args);
                        _conditionalFormatCommands.AddRule(batch, a.SheetName!, a.RangeAddress!, a.RuleType!, a.OperatorType, a.Formula1, a.Formula2, a.InteriorColor, a.InteriorPattern, a.FontColor, a.FontBold, a.FontItalic, a.BorderStyle, a.BorderColor);
                    }),
                    ConditionalFormatAction.ClearRules => ExecuteVoid(() =>
                    {
                        var a = GetArg<ConditionalFormatClearArgs>(request.Args);
                        _conditionalFormatCommands.ClearRules(batch, a.SheetName!, a.RangeAddress!);
                    }),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported conditionalformat action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown conditionalformat action: {action}" };
        });
    }

    private Task<DaemonResponse> HandleVbaCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<VbaAction>(action, out var vbaAction))
            {
                return vbaAction switch
                {
                    VbaAction.List => SerializeResult(_vbaCommands.List(batch)),
                    VbaAction.View => SerializeResult(_vbaCommands.View(batch, GetArg<VbaModuleArgs>(request.Args).ModuleName!)),
                    VbaAction.Import => ExecuteVoid(() =>
                    {
                        var a = GetArg<VbaImportArgs>(request.Args);
                        _vbaCommands.Import(batch, a.ModuleName!, a.VbaCode!);
                    }),
                    VbaAction.Update => ExecuteVoid(() =>
                    {
                        var a = GetArg<VbaImportArgs>(request.Args);
                        _vbaCommands.Update(batch, a.ModuleName!, a.VbaCode!);
                    }),
                    VbaAction.Run => ExecuteVoid(() =>
                    {
                        var a = GetArg<VbaRunArgs>(request.Args);
                        var timeout = a.TimeoutSeconds.HasValue ? TimeSpan.FromSeconds(a.TimeoutSeconds.Value) : (TimeSpan?)null;
                        _vbaCommands.Run(batch, a.ProcedureName!, timeout, a.Parameters?.ToArray() ?? []);
                    }),
                    VbaAction.Delete => ExecuteVoid(() => _vbaCommands.Delete(batch, GetArg<VbaModuleArgs>(request.Args).ModuleName!)),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported vba action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown vba action: {action}" };
        });
    }

    private Task<DaemonResponse> HandleDataModelCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<DataModelAction>(action, out var modelAction))
            {
                return modelAction switch
                {
                    DataModelAction.ListTables => SerializeResult(_dataModelCommands.ListTables(batch)),
                    DataModelAction.ListColumns => SerializeResult(_dataModelCommands.ListColumns(batch, GetArg<DataModelTableArgs>(request.Args).TableName!)),
                    DataModelAction.ReadTable => SerializeResult(_dataModelCommands.ReadTable(batch, GetArg<DataModelTableArgs>(request.Args).TableName!)),
                    DataModelAction.ReadInfo => SerializeResult(_dataModelCommands.ReadInfo(batch)),
                    DataModelAction.ListMeasures => SerializeResult(_dataModelCommands.ListMeasures(batch, GetArg<DataModelTableArgs>(request.Args).TableName)),
                    DataModelAction.Read => SerializeResult(_dataModelCommands.Read(batch, GetArg<DataModelMeasureArgs>(request.Args).MeasureName!)),
                    DataModelAction.CreateMeasure => ExecuteVoid(() =>
                    {
                        var a = GetArg<DataModelCreateMeasureArgs>(request.Args);
                        _dataModelCommands.CreateMeasure(batch, a.TableName!, a.MeasureName!, a.DaxFormula!, a.FormatType, a.Description);
                    }),
                    DataModelAction.UpdateMeasure => ExecuteVoid(() =>
                    {
                        var a = GetArg<DataModelUpdateMeasureArgs>(request.Args);
                        _dataModelCommands.UpdateMeasure(batch, a.MeasureName!, a.DaxFormula, a.FormatType, a.Description);
                    }),
                    DataModelAction.DeleteMeasure => ExecuteVoid(() => _dataModelCommands.DeleteMeasure(batch, GetArg<DataModelMeasureArgs>(request.Args).MeasureName!)),
                    DataModelAction.DeleteTable => ExecuteVoid(() => _dataModelCommands.DeleteTable(batch, GetArg<DataModelTableArgs>(request.Args).TableName!)),
                    DataModelAction.RenameTable => SerializeResult(_dataModelCommands.RenameTable(batch, GetArg<DataModelRenameTableArgs>(request.Args).OldName!, GetArg<DataModelRenameTableArgs>(request.Args).NewName!)),
                    DataModelAction.Refresh => ExecuteVoid(() =>
                    {
                        var a = GetArg<DataModelRefreshArgs>(request.Args);
                        if (a.TimeoutSeconds.HasValue)
                            _dataModelCommands.Refresh(batch, a.TableName, TimeSpan.FromSeconds(a.TimeoutSeconds.Value));
                        else
                            _dataModelCommands.Refresh(batch, a.TableName);
                    }),
                    DataModelAction.Evaluate => SerializeResult(_dataModelCommands.Evaluate(batch, GetArg<DataModelEvaluateArgs>(request.Args).DaxQuery!)),
                    DataModelAction.ExecuteDmv => SerializeResult(_dataModelCommands.ExecuteDmv(batch, GetArg<DataModelDmvArgs>(request.Args).DmvQuery!)),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported datamodel action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown datamodel action: {action}" };
        });
    }

    private Task<DaemonResponse> HandleDataModelRelCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<DataModelRelAction>(action, out var relAction))
            {
                return relAction switch
                {
                    DataModelRelAction.ListRelationships => SerializeResult(_dataModelCommands.ListRelationships(batch)),
                    DataModelRelAction.ReadRelationship => SerializeResult(() =>
                    {
                        var a = GetArg<DataModelRelationshipArgs>(request.Args);
                        return _dataModelCommands.ReadRelationship(batch, a.FromTable!, a.FromColumn!, a.ToTable!, a.ToColumn!);
                    }),
                    DataModelRelAction.CreateRelationship => ExecuteVoid(() =>
                    {
                        var a = GetArg<DataModelCreateRelationshipArgs>(request.Args);
                        _dataModelCommands.CreateRelationship(batch, a.FromTable!, a.FromColumn!, a.ToTable!, a.ToColumn!, a.Active ?? true);
                    }),
                    DataModelRelAction.UpdateRelationship => ExecuteVoid(() =>
                    {
                        var a = GetArg<DataModelUpdateRelationshipArgs>(request.Args);
                        _dataModelCommands.UpdateRelationship(batch, a.FromTable!, a.FromColumn!, a.ToTable!, a.ToColumn!, a.Active ?? true);
                    }),
                    DataModelRelAction.DeleteRelationship => ExecuteVoid(() =>
                    {
                        var a = GetArg<DataModelRelationshipArgs>(request.Args);
                        _dataModelCommands.DeleteRelationship(batch, a.FromTable!, a.FromColumn!, a.ToTable!, a.ToColumn!);
                    }),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported datamodelrel action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown datamodelrel action: {action}" };
        });
    }

    private Task<DaemonResponse> HandleSlicerCommandAsync(string action, DaemonRequest request)
    {
        return WithSessionAsync(request.SessionId, batch =>
        {
            if (TryParseAction<SlicerAction>(action, out var slicerAction))
            {
                return slicerAction switch
                {
                    SlicerAction.ListSlicers => SerializeResult(() =>
                    {
                        var a = GetArg<SlicerListArgs>(request.Args);
                        return _pivotTableCommands.ListSlicers(batch, a.PivotTableName);
                    }),
                    SlicerAction.CreateSlicer => SerializeResult(() =>
                    {
                        var a = GetArg<SlicerFromPivotArgs>(request.Args);
                        return _pivotTableCommands.CreateSlicer(batch, a.PivotTableName!, a.SourceFieldName!, a.SlicerName!, a.DestinationSheet!, BuildSlicerPosition(a.Left, a.Top, a.Width, a.Height));
                    }),
                    SlicerAction.SetSlicerSelection => SerializeResult(() =>
                    {
                        var a = GetArg<SlicerFilterArgs>(request.Args);
                        var items = a.SelectedItems?.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList() ?? [];
                        return _pivotTableCommands.SetSlicerSelection(batch, a.SlicerName!, items, !(a.MultiSelect ?? false));
                    }),
                    SlicerAction.DeleteSlicer => SerializeResult(() =>
                    {
                        var a = GetArg<SlicerArgs>(request.Args);
                        return _pivotTableCommands.DeleteSlicer(batch, a.SlicerName!);
                    }),
                    SlicerAction.ListTableSlicers => SerializeResult(() =>
                    {
                        var a = GetArg<SlicerListArgs>(request.Args);
                        return _tableCommands.ListTableSlicers(batch, a.TableName);
                    }),
                    SlicerAction.CreateTableSlicer => SerializeResult(() =>
                    {
                        var a = GetArg<SlicerFromTableArgs>(request.Args);
                        return _tableCommands.CreateTableSlicer(batch, a.TableName!, a.ColumnName!, a.SlicerName!, a.DestinationSheet!, BuildSlicerPosition(a.Left, a.Top, a.Width, a.Height));
                    }),
                    SlicerAction.SetTableSlicerSelection => SerializeResult(() =>
                    {
                        var a = GetArg<SlicerFilterArgs>(request.Args);
                        var items = a.SelectedItems?.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList() ?? [];
                        return _tableCommands.SetTableSlicerSelection(batch, a.SlicerName!, items, !(a.MultiSelect ?? false));
                    }),
                    SlicerAction.DeleteTableSlicer => SerializeResult(() =>
                    {
                        var a = GetArg<SlicerArgs>(request.Args);
                        return _tableCommands.DeleteTableSlicer(batch, a.SlicerName!);
                    }),
                    _ => new DaemonResponse { Success = false, ErrorMessage = $"Unsupported slicer action: {action}" }
                };
            }

            return new DaemonResponse { Success = false, ErrorMessage = $"Unknown slicer action: {action}" };
        });
    }

    private static string BuildSlicerPosition(double? left, double? top, double? width, double? height)
    {
        // Build position string like "100,100" or "100,100,200,150"
        var parts = new List<string>();
        if (left.HasValue) parts.Add(left.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (top.HasValue) parts.Add(top.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (width.HasValue) parts.Add(width.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (height.HasValue) parts.Add(height.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        return parts.Count > 0 ? string.Join(",", parts) : "100,100";
    }

    // === HELPERS ===

    private static bool TryParseAction<T>(string actionString, out T action) where T : struct, Enum
    {
        var pascalCase = ToPascalCase(actionString);
        return Enum.TryParse(pascalCase, ignoreCase: true, out action);
    }

    private static string ToPascalCase(string kebabCase)
    {
        if (string.IsNullOrWhiteSpace(kebabCase)) return string.Empty;

        var parts = kebabCase.Split('-', StringSplitOptions.RemoveEmptyEntries);
        var builder = new StringBuilder();

        foreach (var part in parts)
        {
            builder.Append(char.ToUpperInvariant(part[0]));
            if (part.Length > 1)
            {
                builder.Append(part[1..]);
            }
        }

        return builder.ToString();
    }

    private Task<DaemonResponse> WithSessionAsync(string? sessionId, Func<IExcelBatch, DaemonResponse> action)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return Task.FromResult(new DaemonResponse { Success = false, ErrorMessage = "sessionId is required" });
        }

        var batch = _sessionManager.GetSession(sessionId);
        if (batch == null)
        {
            return Task.FromResult(new DaemonResponse { Success = false, ErrorMessage = $"Session '{sessionId}' not found" });
        }

        // Check if Excel process is still alive before attempting operation
        if (!batch.IsExcelProcessAlive())
        {
            // Excel died - clean up the dead session
            _sessionManager.CloseSession(sessionId, save: false, force: true);
            return Task.FromResult(new DaemonResponse
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
        catch (Exception ex)
        {
            return Task.FromResult(new DaemonResponse { Success = false, ErrorMessage = ex.Message });
        }
    }

    private static T? DeserializeArgs<T>(string? args) where T : class
    {
        if (string.IsNullOrEmpty(args)) return null;
        return JsonSerializer.Deserialize<T>(args, DaemonProtocol.JsonOptions);
    }

    private static T GetArg<T>(string? args) where T : class, new()
    {
        return DeserializeArgs<T>(args) ?? new T();
    }

    private static DaemonResponse SerializeResult<T>(T result)
    {
        return new DaemonResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(result, DaemonProtocol.JsonOptions)
        };
    }

    private static DaemonResponse SerializeResult<T>(Func<T> action)
    {
        var result = action();
        return new DaemonResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(result, DaemonProtocol.JsonOptions)
        };
    }

    private static DaemonResponse ExecuteVoid(Action action)
    {
        action();
        return new DaemonResponse { Success = true };
    }

    private async Task MonitorIdleTimeoutAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(30), cancellationToken);

                if (_sessionManager.GetActiveSessions().Count == 0)
                {
                    var idleTime = DateTime.UtcNow - _lastActivityTime;
                    if (idleTime >= _idleTimeout)
                    {
                        Console.Error.WriteLine($"Daemon idle for {idleTime.TotalMinutes:F1} minutes, shutting down");
                        _shutdownCts.Cancel();
                        break;
                    }
                }
                else
                {
                    _lastActivityTime = DateTime.UtcNow;
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _shutdownCts.Cancel();

        // Cleanup tray
        if (_tray != null)
        {
            try
            {
                Application.Exit();
                _tray.Dispose();
            }
            catch { }
            _tray = null;
        }

        _sessionManager.Dispose();
        _shutdownCts.Dispose();

        if (_instanceMutex != null)
        {
            try { _instanceMutex.ReleaseMutex(); } catch { }
            _instanceMutex.Dispose();
        }

        DaemonSecurity.DeleteLockFile();
    }
}

// === ARGUMENT TYPES ===

// Session
internal sealed class SessionOpenArgs
{
    public string? FilePath { get; set; }
    public int? TimeoutSeconds { get; set; }
}
internal sealed class SessionCloseArgs { public bool Save { get; set; } }

// Sheet
internal sealed class SheetArgs { public string? SheetName { get; set; } }
internal sealed class SheetRenameArgs { public string? SheetName { get; set; } public string? NewName { get; set; } }
internal sealed class SheetCopyArgs { public string? SourceSheet { get; set; } public string? TargetSheet { get; set; } }
internal sealed class SheetMoveArgs { public string? SheetName { get; set; } public string? BeforeSheet { get; set; } public string? AfterSheet { get; set; } }
internal sealed class SheetCopyToFileArgs { public string? SourceFile { get; set; } public string? SourceSheet { get; set; } public string? TargetFile { get; set; } public string? TargetSheetName { get; set; } public string? BeforeSheet { get; set; } public string? AfterSheet { get; set; } }
internal sealed class SheetMoveToFileArgs { public string? SourceFile { get; set; } public string? SourceSheet { get; set; } public string? TargetFile { get; set; } public string? BeforeSheet { get; set; } public string? AfterSheet { get; set; } }
internal sealed class SheetTabColorArgs { public string? SheetName { get; set; } public int? Red { get; set; } public int? Green { get; set; } public int? Blue { get; set; } }
internal sealed class SheetVisibilityArgs { public string? SheetName { get; set; } public string? Visibility { get; set; } }

// Range
internal sealed class RangeArgs { public string? SheetName { get; set; } public string? Range { get; set; } }
internal sealed class RangeSetValuesArgs { public string? SheetName { get; set; } public string? Range { get; set; } public List<List<object?>>? Values { get; set; } }
internal sealed class RangeSetFormulasArgs { public string? SheetName { get; set; } public string? Range { get; set; } public List<List<string>>? Formulas { get; set; } }
internal sealed class RangeCopyArgs { public string? SourceSheet { get; set; } public string? SourceRange { get; set; } public string? TargetSheet { get; set; } public string? TargetRange { get; set; } }
internal sealed class RangeInsertCellsArgs { public string? SheetName { get; set; } public string? Range { get; set; } public string? ShiftDirection { get; set; } }
internal sealed class RangeDeleteCellsArgs { public string? SheetName { get; set; } public string? Range { get; set; } public string? ShiftDirection { get; set; } }
internal sealed class RangeFindArgs { public string? SheetName { get; set; } public string? Range { get; set; } public string? SearchValue { get; set; } public bool? MatchCase { get; set; } public bool? MatchEntireCell { get; set; } public bool? SearchFormulas { get; set; } public bool? SearchValues { get; set; } public bool? SearchComments { get; set; } }
internal sealed class RangeReplaceArgs { public string? SheetName { get; set; } public string? Range { get; set; } public string? FindValue { get; set; } public string? ReplaceValue { get; set; } public bool? MatchCase { get; set; } public bool? MatchEntireCell { get; set; } public bool? ReplaceAll { get; set; } }
internal sealed class RangeSortArgs { public string? SheetName { get; set; } public string? Range { get; set; } public List<SortColumnArg>? SortColumns { get; set; } public bool? HasHeaders { get; set; } }
internal sealed class SortColumnArg { public int ColumnIndex { get; set; } public bool Ascending { get; set; } = true; }
internal sealed class RangeCellArgs { public string? SheetName { get; set; } public string? CellAddress { get; set; } }
internal sealed class RangeHyperlinkArgs { public string? SheetName { get; set; } public string? CellAddress { get; set; } public string? Url { get; set; } public string? DisplayText { get; set; } public string? Tooltip { get; set; } }
internal sealed class RangeNumberFormatArgs { public string? SheetName { get; set; } public string? Range { get; set; } public string? FormatCode { get; set; } }
internal sealed class RangeNumberFormatsArgs { public string? SheetName { get; set; } public string? Range { get; set; } public List<List<string>>? Formats { get; set; } }
internal sealed class RangeStyleArgs { public string? SheetName { get; set; } public string? Range { get; set; } public string? StyleName { get; set; } }
internal sealed class RangeFormatArgs { public string? SheetName { get; set; } public string? Range { get; set; } public string? FontName { get; set; } public double? FontSize { get; set; } public bool? Bold { get; set; } public bool? Italic { get; set; } public bool? Underline { get; set; } public string? FontColor { get; set; } public string? FillColor { get; set; } public string? BorderStyle { get; set; } public string? BorderColor { get; set; } public string? BorderWeight { get; set; } public string? HorizontalAlignment { get; set; } public string? VerticalAlignment { get; set; } public bool? WrapText { get; set; } public int? Orientation { get; set; } }
internal sealed class RangeValidationArgs { public string? SheetName { get; set; } public string? Range { get; set; } public string? ValidationType { get; set; } public string? ValidationOperator { get; set; } public string? Formula1 { get; set; } public string? Formula2 { get; set; } public bool? ShowInputMessage { get; set; } public string? InputTitle { get; set; } public string? InputMessage { get; set; } public bool? ShowErrorAlert { get; set; } public string? ErrorStyle { get; set; } public string? ErrorTitle { get; set; } public string? ErrorMessage { get; set; } public bool? IgnoreBlank { get; set; } public bool? ShowDropdown { get; set; } }
internal sealed class RangeLockArgs { public string? SheetName { get; set; } public string? Range { get; set; } public bool? Locked { get; set; } }

// Table
internal sealed class TableArgs { public string? TableName { get; set; } }
internal sealed class TableCreateArgs { public string? SheetName { get; set; } public string? TableName { get; set; } public string? Range { get; set; } public bool? HasHeaders { get; set; } public string? TableStyle { get; set; } }
internal sealed class TableRenameArgs { public string? TableName { get; set; } public string? NewName { get; set; } }
internal sealed class TableResizeArgs { public string? TableName { get; set; } public string? NewRange { get; set; } }
internal sealed class TableStyleArgs { public string? TableName { get; set; } public string? TableStyle { get; set; } }
internal sealed class TableToggleTotalsArgs { public string? TableName { get; set; } public bool? ShowTotals { get; set; } }
internal sealed class TableColumnTotalArgs { public string? TableName { get; set; } public string? ColumnName { get; set; } public string? TotalFunction { get; set; } }
internal sealed class TableAppendArgs { public string? TableName { get; set; } public List<List<object?>>? Rows { get; set; } }
internal sealed class TableDataArgs { public string? TableName { get; set; } public bool? VisibleOnly { get; set; } }
internal sealed class TableFilterArgs { public string? TableName { get; set; } public string? ColumnName { get; set; } public string? Criteria { get; set; } public List<string>? Values { get; set; } }
internal sealed class TableAddColumnArgs { public string? TableName { get; set; } public string? ColumnName { get; set; } public int? Position { get; set; } }
internal sealed class TableColumnArgs { public string? TableName { get; set; } public string? ColumnName { get; set; } }
internal sealed class TableRenameColumnArgs { public string? TableName { get; set; } public string? OldName { get; set; } public string? NewName { get; set; } }
internal sealed class TableSortArgs { public string? TableName { get; set; } public string? ColumnName { get; set; } public bool? Ascending { get; set; } }
internal sealed class TableSortMultiArgs { public string? TableName { get; set; } public string? SortColumnsJson { get; set; } }
internal sealed class TableColumnFormatArgs { public string? TableName { get; set; } public string? ColumnName { get; set; } public string? FormatCode { get; set; } }
internal sealed class TableDaxArgs { public string? SheetName { get; set; } public string? TableName { get; set; } public string? DaxQuery { get; set; } public string? TargetCell { get; set; } }
internal sealed class TableUpdateDaxArgs { public string? TableName { get; set; } public string? DaxQuery { get; set; } }

// PowerQuery
internal sealed class PowerQueryArgs { public string? QueryName { get; set; } }
internal sealed class PowerQueryCreateArgs { public string? QueryName { get; set; } public string? MCode { get; set; } public string? LoadDestination { get; set; } public string? TargetSheet { get; set; } public string? TargetCellAddress { get; set; } }
internal sealed class PowerQueryUpdateArgs { public string? QueryName { get; set; } public string? MCode { get; set; } public bool? Refresh { get; set; } }
internal sealed class PowerQueryRenameArgs { public string? OldName { get; set; } public string? NewName { get; set; } }
internal sealed class PowerQueryLoadToArgs { public string? QueryName { get; set; } public string? LoadDestination { get; set; } public string? TargetSheet { get; set; } public string? TargetCellAddress { get; set; } }
internal sealed class PowerQueryEvaluateArgs { public string? MCode { get; set; } }

// PivotTable
internal sealed class PivotTableArgs { public string? PivotTableName { get; set; } }
internal sealed class PivotTableRefreshArgs { public string? PivotTableName { get; set; } public TimeSpan? Timeout { get; set; } }
internal sealed class PivotTableFromRangeArgs { public string? SourceSheet { get; set; } public string? SourceRange { get; set; } public string? DestinationSheet { get; set; } public string? DestinationCell { get; set; } public string? PivotTableName { get; set; } }
internal sealed class PivotTableFromTableArgs { public string? TableName { get; set; } public string? DestinationSheet { get; set; } public string? DestinationCell { get; set; } public string? PivotTableName { get; set; } }
internal sealed class PivotTableFromDataModelArgs { public string? TableName { get; set; } public string? DestinationSheet { get; set; } public string? DestinationCell { get; set; } public string? PivotTableName { get; set; } }
internal sealed class PivotFieldArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public int? Position { get; set; } }
internal sealed class PivotValueFieldArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public string? AggregationFunction { get; set; } public string? CustomName { get; set; } }
internal sealed class PivotFieldFilterArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public List<string>? SelectedValues { get; set; } }
internal sealed class PivotFieldSortArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public bool? Ascending { get; set; } }
internal sealed class PivotCalculatedFieldArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public string? Formula { get; set; } }
internal sealed class PivotLayoutArgs { public string? PivotTableName { get; set; } public int? LayoutType { get; set; } }
internal sealed class PivotSubtotalsArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public bool? ShowSubtotals { get; set; } }
internal sealed class PivotGrandTotalsArgs { public string? PivotTableName { get; set; } public bool? ShowRowGrandTotals { get; set; } public bool? ShowColumnGrandTotals { get; set; } }
internal sealed class PivotFieldFunctionArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public string? AggregationFunction { get; set; } }
internal sealed class PivotFieldNameArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public string? CustomName { get; set; } }
internal sealed class PivotFieldFormatArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public string? NumberFormat { get; set; } }
internal sealed class PivotGroupByDateArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public string? Interval { get; set; } }
internal sealed class PivotGroupByNumericArgs { public string? PivotTableName { get; set; } public string? FieldName { get; set; } public double? Start { get; set; } public double? End { get; set; } public double? IntervalSize { get; set; } }
internal sealed class PivotCalculatedMemberArgs { public string? PivotTableName { get; set; } public string? MemberName { get; set; } public string? Formula { get; set; } public string? MemberType { get; set; } public int? SolveOrder { get; set; } public string? DisplayFolder { get; set; } public string? NumberFormat { get; set; } }
internal sealed class TableStructuredRefArgs { public string? TableName { get; set; } public string? Region { get; set; } public string? ColumnName { get; set; } }

// Chart
internal sealed class ChartArgs { public string? ChartName { get; set; } }
internal sealed class ChartFromRangeArgs { public string? SheetName { get; set; } public string? SourceRange { get; set; } public string? ChartType { get; set; } public double? Left { get; set; } public double? Top { get; set; } public double? Width { get; set; } public double? Height { get; set; } public string? ChartName { get; set; } }
internal sealed class ChartFromPivotArgs { public string? PivotTableName { get; set; } public string? SheetName { get; set; } public string? ChartType { get; set; } public double? Left { get; set; } public double? Top { get; set; } public double? Width { get; set; } public double? Height { get; set; } public string? ChartName { get; set; } }
internal sealed class ChartFromTableArgs { public string? TableName { get; set; } public string? SheetName { get; set; } public string? ChartType { get; set; } public double? Left { get; set; } public double? Top { get; set; } public double? Width { get; set; } public double? Height { get; set; } public string? ChartName { get; set; } }
internal sealed class ChartMoveArgs { public string? ChartName { get; set; } public double? Left { get; set; } public double? Top { get; set; } public double? Width { get; set; } public double? Height { get; set; } }
internal sealed class ChartFitArgs { public string? ChartName { get; set; } public string? SheetName { get; set; } public string? RangeAddress { get; set; } }
internal sealed class ChartSourceRangeArgs { public string? ChartName { get; set; } public string? SourceRange { get; set; } }
internal sealed class ChartAddSeriesArgs { public string? ChartName { get; set; } public string? SeriesName { get; set; } public string? ValuesRange { get; set; } public string? CategoryRange { get; set; } }
internal sealed class ChartRemoveSeriesArgs { public string? ChartName { get; set; } public int? SeriesIndex { get; set; } }
internal sealed class ChartTypeArgs { public string? ChartName { get; set; } public string? ChartType { get; set; } }
internal sealed class ChartTitleArgs { public string? ChartName { get; set; } public string? Title { get; set; } }
internal sealed class ChartAxisTitleArgs { public string? ChartName { get; set; } public string? Axis { get; set; } public string? Title { get; set; } }
internal sealed class ChartAxisArgs { public string? ChartName { get; set; } public string? Axis { get; set; } }
internal sealed class ChartAxisFormatArgs { public string? ChartName { get; set; } public string? Axis { get; set; } public string? NumberFormat { get; set; } }
internal sealed class ChartLegendArgs { public string? ChartName { get; set; } public bool? Visible { get; set; } public string? LegendPosition { get; set; } }
internal sealed class ChartStyleArgs { public string? ChartName { get; set; } public int? StyleId { get; set; } }
internal sealed class ChartDataLabelsArgs { public string? ChartName { get; set; } public bool? ShowValue { get; set; } public bool? ShowPercentage { get; set; } public bool? ShowSeriesName { get; set; } public bool? ShowCategoryName { get; set; } public string? Separator { get; set; } public string? LabelPosition { get; set; } public int? SeriesIndex { get; set; } }
internal sealed class ChartAxisScaleArgs { public string? ChartName { get; set; } public string? Axis { get; set; } public double? MinimumScale { get; set; } public double? MaximumScale { get; set; } public double? MajorUnit { get; set; } public double? MinorUnit { get; set; } }
internal sealed class ChartGridlinesArgs { public string? ChartName { get; set; } public string? Axis { get; set; } public bool? ShowMajor { get; set; } public bool? ShowMinor { get; set; } }
internal sealed class ChartSeriesFormatArgs { public string? ChartName { get; set; } public int? SeriesIndex { get; set; } public string? MarkerStyle { get; set; } public int? MarkerSize { get; set; } public string? MarkerBackgroundColor { get; set; } public string? MarkerForegroundColor { get; set; } }
internal sealed class ChartSeriesArgs { public string? ChartName { get; set; } public int? SeriesIndex { get; set; } }
internal sealed class ChartAddTrendlineArgs { public string? ChartName { get; set; } public int? SeriesIndex { get; set; } public string? TrendlineType { get; set; } public bool? DisplayEquation { get; set; } public bool? DisplayRSquared { get; set; } public string? TrendlineName { get; set; } }
internal sealed class ChartDeleteTrendlineArgs { public string? ChartName { get; set; } public int? SeriesIndex { get; set; } public int? TrendlineIndex { get; set; } }
internal sealed class ChartSetTrendlineArgs { public string? ChartName { get; set; } public int? SeriesIndex { get; set; } public int? TrendlineIndex { get; set; } public bool? DisplayEquation { get; set; } public bool? DisplayRSquared { get; set; } public string? TrendlineName { get; set; } }
internal sealed class ChartPlacementArgs { public string? ChartName { get; set; } public int? Placement { get; set; } }

// Connection
internal sealed class ConnectionArgs { public string? ConnectionName { get; set; } }
internal sealed class ConnectionCreateArgs { public string? ConnectionName { get; set; } public string? ConnectionString { get; set; } public string? CommandText { get; set; } public string? Description { get; set; } }
internal sealed class ConnectionRefreshArgs { public string? ConnectionName { get; set; } public int? TimeoutSeconds { get; set; } }
internal sealed class ConnectionLoadToArgs { public string? ConnectionName { get; set; } public string? SheetName { get; set; } }
internal sealed class ConnectionSetPropertiesArgs { public string? ConnectionName { get; set; } public string? ConnectionString { get; set; } public string? CommandText { get; set; } public string? Description { get; set; } public bool? BackgroundQuery { get; set; } public bool? RefreshOnFileOpen { get; set; } public bool? SavePassword { get; set; } public int? RefreshPeriod { get; set; } }

// NamedRange
internal sealed class NamedRangeArgs { public string? ParamName { get; set; } }
internal sealed class NamedRangeWriteArgs { public string? ParamName { get; set; } public string? Value { get; set; } }
internal sealed class NamedRangeCreateArgs { public string? ParamName { get; set; } public string? Reference { get; set; } }

// ConditionalFormat
internal sealed class ConditionalFormatAddArgs { public string? SheetName { get; set; } public string? RangeAddress { get; set; } public string? RuleType { get; set; } public string? OperatorType { get; set; } public string? Formula1 { get; set; } public string? Formula2 { get; set; } public string? InteriorColor { get; set; } public string? InteriorPattern { get; set; } public string? FontColor { get; set; } public bool? FontBold { get; set; } public bool? FontItalic { get; set; } public string? BorderStyle { get; set; } public string? BorderColor { get; set; } }
internal sealed class ConditionalFormatClearArgs { public string? SheetName { get; set; } public string? RangeAddress { get; set; } }

// VBA
internal sealed class VbaModuleArgs { public string? ModuleName { get; set; } }
internal sealed class VbaImportArgs { public string? ModuleName { get; set; } public string? VbaCode { get; set; } }
internal sealed class VbaRunArgs { public string? ProcedureName { get; set; } public int? TimeoutSeconds { get; set; } public List<string>? Parameters { get; set; } }

// DataModel
internal sealed class DataModelTableArgs { public string? TableName { get; set; } }
internal sealed class DataModelMeasureArgs { public string? MeasureName { get; set; } }
internal sealed class DataModelCreateMeasureArgs { public string? TableName { get; set; } public string? MeasureName { get; set; } public string? DaxFormula { get; set; } public string? FormatType { get; set; } public string? Description { get; set; } }
internal sealed class DataModelUpdateMeasureArgs { public string? MeasureName { get; set; } public string? DaxFormula { get; set; } public string? FormatType { get; set; } public string? Description { get; set; } }
internal sealed class DataModelRelationshipArgs { public string? FromTable { get; set; } public string? FromColumn { get; set; } public string? ToTable { get; set; } public string? ToColumn { get; set; } }
internal sealed class DataModelCreateRelationshipArgs { public string? FromTable { get; set; } public string? FromColumn { get; set; } public string? ToTable { get; set; } public string? ToColumn { get; set; } public bool? Active { get; set; } }
internal sealed class DataModelUpdateRelationshipArgs { public string? FromTable { get; set; } public string? FromColumn { get; set; } public string? ToTable { get; set; } public string? ToColumn { get; set; } public bool? Active { get; set; } }
internal sealed class DataModelRenameTableArgs { public string? OldName { get; set; } public string? NewName { get; set; } }
internal sealed class DataModelRefreshArgs { public string? TableName { get; set; } public int? TimeoutSeconds { get; set; } }
internal sealed class DataModelEvaluateArgs { public string? DaxQuery { get; set; } }
internal sealed class DataModelDmvArgs { public string? DmvQuery { get; set; } }

// Slicer
internal sealed class SlicerArgs { public string? SlicerName { get; set; } }
internal sealed class SlicerListArgs { public string? PivotTableName { get; set; } public string? TableName { get; set; } public string? SheetName { get; set; } }
internal sealed class SlicerPositionArgs { public double? Top { get; set; } public double? Left { get; set; } public double? Width { get; set; } public double? Height { get; set; } }
internal sealed class SlicerFromPivotArgs { public string? PivotTableName { get; set; } public string? SourceFieldName { get; set; } public string? SlicerName { get; set; } public string? DestinationSheet { get; set; } public double? Top { get; set; } public double? Left { get; set; } public double? Width { get; set; } public double? Height { get; set; } }
internal sealed class SlicerFromTableArgs { public string? TableName { get; set; } public string? ColumnName { get; set; } public string? SlicerName { get; set; } public string? DestinationSheet { get; set; } public double? Top { get; set; } public double? Left { get; set; } public double? Width { get; set; } public double? Height { get; set; } }
internal sealed class SlicerFilterArgs { public string? SlicerName { get; set; } public string? SelectedItems { get; set; } public bool? MultiSelect { get; set; } }
