using System.Globalization;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Nerdbank.Streams;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Rpc;
using StreamJsonRpc;

namespace Sbroenne.ExcelMcp.Service.Workers;

internal sealed class ExcelWorkerHost : IDisposable
{
    internal const string Mode = "--excelmcp-private-worker";
    private readonly ExcelMcpService _service;
    private readonly WorkerRegistry _registry = new();
    private readonly Lock _stateLock = new();
    private readonly Lock _snapshotLock = new();
    private long _snapshotRevision;
    private readonly SemaphoreSlim _execution = new(1, 1);
    private readonly CancellationTokenSource _stop = new();
    private readonly DaemonHost _server;
    private WorkerRegistration _registration;
    private WorkerSnapshot _snapshot;
    private string? _pendingOperation;
    private string? _pendingCommand;
    private bool _releasing;
    private Task? _refresh;
    private string? _statusError;

    internal Action? AfterOperationStatusCaptureForTests { get; set; }

    internal ExcelWorkerHost(WorkerRegistration registration, ExcelMcpService? service = null)
    {
        _service = service ?? new ExcelMcpService();
        _registration = registration;
        _snapshot = new(registration, [], EmptySessionList(), null, null,
            new Dictionary<string, IReadOnlyList<string>>());
        _server = new DaemonHost(registration.PipeName, pipe =>
            new JsonRpc(pipe, pipe.ReadSlice(DaemonHost.MaximumConnectionInputBytes),
                new Target(this, WorkerPipeIdentity.GetClient(pipe.SafePipeHandle)))
            {
                CancelLocallyInvokedMethodsWhenConnectionIsClosed = true
            }, maximumConnections: 8);
    }

    internal static bool IsWorkerMode(string[] args) =>
        args.Length > 0 && args[0] == Mode;

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "The private worker process boundary reports bootstrap failures to stderr and exits nonzero; it never emits product or MCP protocol output.")]
    internal static async Task<int> RunAsync(string[] args)
    {
        try
        {
            if (args.Length != 8)
                throw new ArgumentException("Invalid private Excel worker startup arguments.");
            var controller = new WorkerProcess(
                int.Parse(args[2], CultureInfo.InvariantCulture),
                long.Parse(args[3], CultureInfo.InvariantCulture));
            if (!controller.IsAlive() || args[5] != WorkerRegistry.Build)
                throw new InvalidOperationException("The Excel controller identity or build could not be confirmed.");
            ExcelProcessIdentity? excel = args[6] == "0" ? null : new(
                int.Parse(args[6], CultureInfo.InvariantCulture),
                long.Parse(args[7], CultureInfo.InvariantCulture));
            using var host = new ExcelWorkerHost(new(
                args[1], WorkerProcess.Current, controller, args[4], WorkerRegistry.Build, excel));
            await host.RunAsync();
            return 0;
        }
        catch (Exception ex)
        {
            await Console.Error.WriteLineAsync($"Excel worker failed: {ex}");
            return 1;
        }
    }

    internal async Task RunAsync()
    {
        _registry.Write(_registration);
        SessionManager.ExcelProcessIdentityTracked += RecordExcel;
        SessionManager.ExcelProcessIdentityPreserved += RecordPreservation;
        SessionManager.ExcelProcessIdentityReleased += RecordRelease;
        using var application = new HostBuilder()
            .ConfigureServices(services => services.AddSingleton<IHostedService>(_server))
            .Build();
        try
        {
            await application.StartAsync();
            var monitor = MonitorControllerAsync();
            await _server.ExecuteTask!;
            await _stop.CancelAsync();
            await monitor;
            await application.StopAsync();
        }
        finally
        {
            SessionManager.ExcelProcessIdentityTracked -= RecordExcel;
            SessionManager.ExcelProcessIdentityPreserved -= RecordPreservation;
            SessionManager.ExcelProcessIdentityReleased -= RecordRelease;
        }
        if (!_registration.ReleaseConfirmed)
            _registry.Remove(_registration.PipeName);
    }

    private void RecordExcel(ExcelProcessIdentity identity)
    {
        lock (_stateLock)
        {
            if (_registration.Excel is { } previous && previous != identity)
                throw new InvalidOperationException("An Excel worker cannot own more than one Excel application.");
            _registration = _registration with { Excel = identity, OwnsExcel = true };
            _registry.Write(_registration);
        }
    }

    private void RecordPreservation(ExcelProcessIdentity identity)
    {
        lock (_stateLock)
        {
            if (_registration.Excel != identity)
                throw new InvalidOperationException("The released Excel application does not match this worker.");
            _registration = _registration with { ExcelPreserved = true };
            _registry.Write(_registration);
        }

    }

    private void RecordRelease(ExcelProcessIdentity identity)
    {
        lock (_stateLock)
        {
            if (_registration.Excel != identity)
                throw new InvalidOperationException("The released Excel application does not match this worker.");
            var released = _registration with
            {
                Excel = null,
                OwnsExcel = false,
                ExcelPreserved = false,
                ReleasedExcel = identity
            };
            _registry.Write(released);
            _registration = released;
        }
    }

    private WorkerSnapshot Claim(WorkerClaim claim, int clientId)
    {
        lock (_stateLock)
        {
            if (clientId != claim.Controller.ProcessId || !claim.Controller.IsAlive()
                || claim.Build != WorkerRegistry.Build)
                throw new InvalidOperationException("The connecting Excel controller identity could not be confirmed.");
            if ((_registration.Controller != claim.Controller || _registration.Generation != claim.Generation)
                && _registration.Controller.IsAlive())
                throw new InvalidOperationException("This Excel worker still belongs to another running controller.");
            _registration = _registration with { Controller = claim.Controller, Generation = claim.Generation };
            _registry.Write(_registration);
            return SnapshotCore();
        }
    }

    private void Authorize(string generation, int clientId)
    {
        lock (_stateLock)
        {
            if (generation != _registration.Generation || clientId != _registration.Controller.ProcessId
                || !_registration.Controller.IsAlive())
                throw new InvalidOperationException("This request is not from the owning Excel controller.");
        }
    }

    private WorkerSnapshot SnapshotCore() => _snapshot with
    {
        Registration = _registration,
        PendingOperation = _pendingOperation,
        PendingCommand = _pendingCommand,
        StatusError = _statusError,
        Releasing = _releasing,
        Revision = ++_snapshotRevision
    };

    private void RefreshSnapshot()
    {
        lock (_snapshotLock)
        {
            for (var attempt = 0; ; attempt++)
            {
                // Reconcile completed Save As paths without COM work or session cleanup.
                var operations = _service.SessionManager.GetActiveSessions(statusOnly: true)
                    .Select(session => (session.SessionId, Status: _service.SessionManager.GetOperationStatus(session.SessionId)))
                    .ToArray();
                AfterOperationStatusCaptureForTests?.Invoke();
                var sessions = _service.SessionManager.GetActiveSessions(statusOnly: true);
                var list = _service.ProcessAsync(new ServiceRequest { Command = "session.list" }).GetAwaiter().GetResult();
                if (!list.Success)
                    throw new InvalidOperationException($"Could not record Excel worker status: {list.ErrorMessage}");
                if (sessions.Count != operations.Length || operations.Any(observed =>
                    {
                        var current = _service.SessionManager.GetOperationStatus(observed.SessionId, reconcilePaths: false);
                        return current?.StartedAtUtc != observed.Status?.StartedAtUtc
                            || current?.ExecutionPending != observed.Status?.ExecutionPending;
                    }))
                {
                    if (attempt == 1)
                        throw new InvalidOperationException("Excel operations changed while status was being recorded.");
                    continue;
                }
                lock (_stateLock)
                {
                    _snapshot = new(_registration, sessions, list, _pendingOperation, _pendingCommand,
                        OwnershipKeys: sessions.ToDictionary(session => session.SessionId,
                            session => _service.SessionManager.GetOwnedFilePathKeys(session.SessionId)));
                    _statusError = null;
                }
                return;
            }
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A stored-status refresh reports its error in subsequent snapshots without replacing a completed Excel result or claiming that stale state is healthy.")]
    private void TryRefreshSnapshot()
    {
        try
        {
            RefreshSnapshot();
        }
        catch (Exception ex)
        {
            lock (_stateLock)
                _statusError = $"Excel status could not be refreshed: {ex.GetType().Name}: {ex.Message}";
            System.Diagnostics.Trace.TraceError("Excel status could not be refreshed: {0}", ex);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Failure to record completed worker shutdown returns both that error and the exact completed workbook response. It retains the worker and never repeats or rolls back workbook changes.")]
    private async Task<WorkerExecution> ExecuteAsync(WorkerCall call, int clientId, CancellationToken cancellationToken)
    {
        Authorize(call.Generation, clientId);
        using var closeTimings = call.Request.Command == "session.close" ? SessionDiagnostics.BeginCloseTimings() : null;
        if (call.Request.Command is "window.show" or "window.hide"
            or "powerquerytrace.list" or "powerquerytrace.summarize")
        {
            lock (_stateLock)
                if (_releasing)
                    return Completed(Failure(call.Request, "SessionUnavailable",
                        "The Excel worker is being released. No work was started."), accepted: false);
            call.Request.CallerCancellation = cancellationToken;
            return Completed(await Task.Run(() => _service.ProcessAsync(call.Request), CancellationToken.None), accepted: true);
        }
        if (!Guid.TryParseExact(call.OperationId, "N", out _))
            throw new ArgumentException("Invalid Excel operation identifier.");
        if (!await _execution.WaitAsync(0, cancellationToken))
            return Completed(Failure(call.Request, "Busy",
                "Another request is still running in this Excel application. No new work was started."), accepted: false);
        try
        {
            lock (_stateLock)
            {
                if (_releasing)
                    return Completed(Failure(call.Request, "SessionUnavailable",
                        "The Excel worker is being released. No work was started."), accepted: false);
                if (_service.SessionManager.HasPendingExecution)
                    return Completed(Failure(call.Request, "Busy",
                        "An earlier Excel operation is still pending. No new work was started."), accepted: false);
                _pendingOperation = call.OperationId;
                _pendingCommand = call.Request.Command;
            }
            call.Request.CallerCancellation = cancellationToken;
            using var abandoned = cancellationToken.Register(() => _service.RevealAbandonedRequest(call.Request, null));
            var response = await Task.Run(() => ProcessAsync(call.Request), CancellationToken.None);
            if (cancellationToken.IsCancellationRequested)
                _service.RevealAbandonedRequest(call.Request, response);
            lock (_stateLock)
            {
                _pendingOperation = null;
                _pendingCommand = null;
            }
            using (SessionDiagnostics.BeginCloseStage("worker-completion"))
            {
                try
                {
                    if (response.Success && call.Request.Command == "session.close"
                        && !_service.SessionManager.HasActiveConnections
                        && !_service.SessionManager.HasPendingExecution
                        && (!_registration.OwnsExcel || _registration.ExcelPreserved))
                        CompleteRelease();
                    else
                        TryRefreshSnapshot();
                }
                catch (Exception ex)
                {
                    lock (_stateLock)
                        _statusError = $"The completed request could not be recorded: {ex.GetType().Name}: {ex.Message}";
                    response = response.WithCompletionFailure(call.Request, ex);
                }
            }
            return Completed(closeTimings is null ? response : response.WithCloseTimings(closeTimings.Snapshot()), accepted: true);
        }

        finally
        {
            lock (_stateLock)
            {
                _pendingOperation = null;
                _pendingCommand = null;
            }
            _execution.Release();
        }
    }

    private WorkerExecution Completed(ServiceResponse response, bool accepted)
    {
        lock (_stateLock)
            return new(response, accepted, SnapshotCore());
    }

    private void CompleteRelease()
    {
        lock (_snapshotLock)
            lock (_stateLock)
            {
                var released = _registration with { ReleaseConfirmed = true };
                _registry.Write(released);
                _registration = released;
                _releasing = true;
                _statusError = null;
                _snapshot = new(_registration, [], EmptySessionList(), null, null,
                    new Dictionary<string, IReadOnlyList<string>>());
            }
        _server.RequestShutdownAfterResponse();
    }

    private Task<ServiceResponse> ProcessAsync(ServiceRequest request)
    {
        if (request.Command == "session.discover" && _registration.Excel is { } excel)
            return DiscoverAsync(excel);
        if (request.Command is "service.shutdown" or "service.status" or "service.ping")
            return Task.FromResult(Failure(request, "InvalidInput", "Service control belongs to the controller, not an Excel worker."));
        if (request.Command is "session.open" or "session.create"
            && (_registration.Excel is not null || _service.SessionCount != 0))
            return Task.FromResult(Failure(request, "Conflict", "This worker already has an Excel application."));
        return _service.ProcessAsync(request);
    }

    private async Task<ServiceResponse> DiscoverAsync(ExcelProcessIdentity excel)
    {
        var discovery = await _service.SessionManager.DiscoverOpenWorkbooksAsync(excel);
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

    private WorkerRelease Released(ServiceResponse response)
    {
        lock (_stateLock)
            return new(response, SnapshotCore());
    }

    private async Task<WorkerRelease> ReleaseAsync(string generation, int clientId)
    {
        Authorize(generation, clientId);
        if (!await _execution.WaitAsync(0))
            return Released(Failure(new ServiceRequest { Command = "service.shutdown" }, "Busy",
                "An Excel request is still running. Its worker was retained without saving or discarding work."));
        try
        {
            TryRefreshSnapshot();
            lock (_stateLock)
            {
                if (_statusError is not null
                    || _snapshot.Sessions.Any(session =>
                        _service.SessionManager.GetOperationStatus(session.SessionId, reconcilePaths: false)?.ExecutionPending == true))
                    return Released(Failure(new ServiceRequest { Command = "service.shutdown" }, "Busy",
                        _statusError ?? "Excel is still executing. Its worker remains available for explicit reconnection."));
                _releasing = true;
            }
            var report = await Task.Run(_service.RequestShutdown);
            if (report is null && !_service.SessionManager.HasActiveConnections)
                CompleteRelease();
            else
                _server.RequestShutdownAfterResponse();
            return Released(new ServiceResponse { Success = true, PartialState = report });
        }
        finally
        {
            _execution.Release();
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "A failed native-window reveal is logged and retried without saving, closing, releasing, or terminating retained work. Controller loss does not turn a reveal failure into worker shutdown.")]
    private async Task MonitorControllerAsync()
    {
        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(2));
        var revealed = new HashSet<string>(StringComparer.Ordinal);
        try
        {
            while (await timer.WaitForNextTickAsync(_stop.Token))
            {
                lock (_stateLock)
                {
                    if (_registration.Controller.IsAlive())
                    {
                        revealed.Clear();
                        continue;
                    }
                    if (_releasing)
                        continue;
                }
                foreach (var sessionId in _service.SessionManager.ActiveSessionIds.Where(id => !revealed.Contains(id)))
                {
                    try
                    {
                        _service.SessionManager.RevealForRecovery(sessionId);
                        revealed.Add(sessionId);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Trace.TraceError("Could not reveal Excel after its controller exited: {0}", ex);
                    }
                }
                // Empty workers have no workbook ownership to hand over.
                if (_service.SessionCount == 0 && await _execution.WaitAsync(0, _stop.Token))
                {
                    try
                    {
                        lock (_stateLock)
                        {
                            if (_registration.Controller.IsAlive())
                                continue;
                            _releasing = true;
                        }
                        await Task.Run(_service.RequestShutdown);
                        _server.RequestShutdown();
                        return;
                    }
                    finally
                    {
                        _execution.Release();
                    }
                }
            }
        }
        catch (OperationCanceledException) when (_stop.IsCancellationRequested)
        {
            return;
        }
    }

    internal static ServiceResponse Failure(ServiceRequest request, string category, string message) => new()
    {
        Success = false,
        Command = request.Command,
        SessionId = request.SessionId,
        ErrorCategory = category,
        ErrorMessage = message
    };

    private static ServiceResponse EmptySessionList() => new()
    {
        Success = true,
        Result = """{"success":true,"sessions":[],"count":0}"""
    };

    public void Dispose()
    {
        _service.Dispose();
        _stop.Dispose();
        _execution.Dispose();
    }

    private sealed class Target(ExcelWorkerHost host, int clientId) : IExcelWorkerRpc
    {
        public Task<WorkerSnapshot> ClaimAsync(WorkerClaim claim, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(host.Claim(claim, clientId));
        }

        public Task<WorkerSnapshot> SnapshotAsync(string generation, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            host.Authorize(generation, clientId);
            lock (host._stateLock)
            {
                if (!host._releasing && (host._refresh is null || host._refresh.IsCompleted))
                    host._refresh = Task.Run(host.TryRefreshSnapshot);
                return Task.FromResult(host.SnapshotCore());
            }
        }

        public Task<WorkerExecution> ExecuteAsync(WorkerCall call, CancellationToken cancellationToken = default) =>
            host.ExecuteAsync(call, clientId, cancellationToken);

        public Task<WorkerRelease> ReleaseAsync(string generation, CancellationToken cancellationToken = default) =>
            host.ReleaseAsync(generation, clientId);
    }
}
