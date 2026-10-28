using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Service.Workers;

namespace Sbroenne.ExcelMcp.Service;

/// <summary>Host-owned controller. Excel objects and synchronous workbook work live only in private workers.</summary>
internal sealed class ExcelController : IDisposable
{
    private readonly ProductProcessLauncher _launcher;
    private readonly WorkerRegistry _registry;
    private readonly Func<ExcelWorkerClient, Task> _initializeWorker;
    private readonly WorkerProcess _identity = WorkerProcess.Current;
    private readonly string _generation = Guid.NewGuid().ToString("N");
    private readonly DateTime _startTime = DateTime.UtcNow;
    private readonly ConcurrentDictionary<string, ExcelWorkerClient> _workers = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, ExcelWorkerClient> _sessions = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, (ExcelWorkerClient Worker, OpenWorkbookInfo Workbook)> _selections = new(StringComparer.Ordinal);
    private sealed record PathReservation(string Token, string? SessionId = null);
    private readonly ConcurrentDictionary<string, PathReservation> _paths = new(StringComparer.OrdinalIgnoreCase);
    private readonly SemaphoreSlim _discovery = new(1, 1);
    private readonly ConcurrentDictionary<string, byte> _newWorkers = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, byte> _abandonedWorkers = new(StringComparer.Ordinal);
    private enum SessionEnd { Closed, WorkerExited }
    private readonly ConcurrentDictionary<string, SessionEnd> _closedSessions = new(StringComparer.Ordinal);
    private readonly CancellationTokenSource _stop = new();
    private readonly DaemonHost _host;
    private readonly DaemonHost _controlHost;
    private readonly Lock _shutdownLock = new();
    private Task<string?>? _shutdown;
    private readonly Task _poll;
    private readonly Lock _utilityLock = new();
    private Task<ExcelWorkerClient>? _utility;
    private bool _disposed;
    internal event Action<ExcelProcessIdentity>? ExcelProcessObserved;
    internal event Action<ExcelProcessIdentity>? ExcelProcessPreserved;

    internal ExcelController(
        Func<ServiceRequest, Func<Task<ServiceResponse>>, Task<ServiceResponse>>? trackRequest = null,
        ProductProcessLauncher? launcher = null, WorkerRegistry? registry = null,
        Func<ExcelWorkerClient, Task>? initializeWorker = null)
    {
        _launcher = launcher ?? ProductProcessLauncher.Current();
        _registry = registry ?? new WorkerRegistry();
        _initializeWorker = initializeWorker ?? (worker => worker.InitializeAsync());
        Task<ServiceResponse> ProcessTrackedAsync(ServiceRequest request) => trackRequest is null
            ? ProcessAsync(request) : trackRequest(request, () => ProcessAsync(request));
        _host = new DaemonHost(ProcessTrackedAsync);
        _controlHost = new DaemonHost(request => ServiceClient.IsControlCommand(request.Command)
            ? ProcessTrackedAsync(request)
            : Task.FromResult(ExcelWorkerHost.Failure(request, "InvalidInput", "This connection accepts service controls only.")));
        _poll = PollAsync();
    }

    internal static bool IsWorkerMode(string[] args) => ExcelWorkerHost.IsWorkerMode(args);
    internal static Task<int> RunWorkerAsync(string[] args) => ExcelWorkerHost.RunAsync(args);
    internal Task RunAsync(string pipeName) => Task.WhenAll(
        _host.RunAsync(pipeName), _controlHost.RunAsync(ServiceClient.GetControlPipeName(pipeName)));

    internal IReadOnlyList<SessionDescriptor> Sessions => _sessions
        .SelectMany(entry => entry.Value.Snapshot?.Sessions.Where(session => session.SessionId == entry.Key) ?? [])
        .DistinctBy(session => session.SessionId).ToArray();

    internal IReadOnlyCollection<ExcelProcessIdentity> ExcelProcesses =>
        _workers.Values.Where(worker => worker.Registration.OwnsExcel)
            .Select(worker => worker.Registration.Excel).OfType<ExcelProcessIdentity>().Distinct().ToArray();

    internal async Task<ServiceResponse> ProcessAsync(ServiceRequest request) =>
        ExcelMcpService.AttachRequestContext(request, await ProcessCoreAsync(request).ConfigureAwait(false));

    private static bool StartsWorkbookWorker(string command) =>
        command is "session.open" or "session.create" or "session.test" or "sheet.copy-to-file" or "sheet.move-to-file";

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "The controller is a failure-transport boundary. It reports the exact worker/transport cause and uncertain workbook outcome without replaying, saving, or discarding work.")]
    private async Task<ServiceResponse> ProcessCoreAsync(ServiceRequest request)
    {
        List<string> reservedPaths = [];
        var reservation = new PathReservation(Guid.NewGuid().ToString("N"),
            request.Command == "workbook.save-as" ? request.SessionId : null);
        ExcelWorkerClient? worker = null;
        var sent = false;
        try
        {
            request.CallerCancellation.ThrowIfCancellationRequested();
            ServiceRegistry.ValidateCommandArguments(request.Command, request.Args);
            if (request.Command.StartsWith("session.", StringComparison.Ordinal))
                ExcelMcpService.ValidateSessionActionArguments(request.Command["session.".Length..], request.Args);
            if (request.Command == "service.ping")
                return new ServiceResponse { Success = true };
            if (request.Command == "service.status")
                return new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new ServiceStatus
                    {
                        Running = true,
                        ProcessId = Environment.ProcessId,
                        SessionCount = _sessions.Count,
                        StartTime = _startTime
                    }, ServiceProtocol.JsonOptions)
                };
            if (request.Command == "session.list")
                return SessionList();
            if (request.Command == "service.shutdown")
            {
                var report = await ShutdownAsync();
                _host.RequestShutdownAfterResponse();
                _controlHost.RequestShutdownAfterResponse();
                return new ServiceResponse { Success = true, PartialState = report };
            }
            if (_shutdown is not null)
                return ExcelWorkerHost.Failure(request, "SessionUnavailable", "The service is stopping. No new work was started.");
            if (request.Command == "session.discover")
                return await DiscoverAsync(request.CallerCancellation);

            if (StartsWorkbookWorker(request.Command))
            {
                using var arguments = JsonDocument.Parse(request.Args ?? "{}");
                string[] operands = request.Command.StartsWith("sheet.", StringComparison.Ordinal)
                    ? ["sourceFile", "targetFile"] : ["filePath"];
                foreach (var operand in operands)
                {
                    var path = arguments.RootElement.TryGetProperty(operand, out var value) ? value.GetString() : null;
                    if (string.IsNullOrWhiteSpace(path))
                        return ExcelWorkerHost.Failure(request, "InvalidInput", $"{operand} is required");
                    var normalized = WorkbookLocation.GetOwnershipKey(path);
                    if (reservedPaths.Contains(normalized, StringComparer.OrdinalIgnoreCase))
                        continue;
                    if (!_paths.TryAdd(normalized, reservation))
                        return ExcelWorkerHost.Failure(request, "Conflict", "This workbook is already reserved by another session or pending request.");
                    reservedPaths.Add(normalized);
                }
                worker = await StartWorkerAsync(null, request.CallerCancellation);
            }
            else if (request.Command == "session.attach")
            {
                using var arguments = JsonDocument.Parse(request.Args ?? "{}");
                var selection = arguments.RootElement.TryGetProperty("workbookId", out var selected)
                    && selected.ValueKind == JsonValueKind.String ? selected.GetString() : null;
                if (selection is null || !_selections.TryGetValue(selection, out var selectedWorkbook))
                    return ExcelWorkerHost.Failure(request, "InvalidInput",
                        selection is null ? "workbookId from session discover is required."
                            : "Unknown open-workbook selection. Discover open workbooks first.");
                worker = selectedWorkbook.Worker;
                if (selectedWorkbook.Workbook.Location is { } location)
                {
                    var keys = selectedWorkbook.Workbook.SessionId is { } selectedSession
                        && worker.Snapshot is { } snapshot
                        ? RecordedPaths(snapshot, selectedSession).Where(key =>
                            WorkbookLocation.MatchesReportedLocation(key, location)).ToArray()
                        : [WorkbookLocation.GetOwnershipKey(location)];
                    foreach (var key in keys)
                    {
                        if (!_paths.TryGetValue(key, out var existing)
                            || existing.SessionId is null || existing.SessionId != selectedWorkbook.Workbook.SessionId)
                        {
                            if (!_paths.TryAdd(key, reservation))
                                return ExcelWorkerHost.Failure(request, "Conflict", "This workbook is reserved by another session or pending request.");
                            reservedPaths.Add(key);
                        }
                    }
                }
            }
            else if (request.Command.StartsWith("powerquerytrace.", StringComparison.Ordinal)
                && JsonNode.Parse(request.Args ?? "{}")?["workbookSessionId"] is JsonValue workbookFilter
                && workbookFilter.TryGetValue<string>(out var filteredSession)
                && _sessions.TryGetValue(filteredSession, out var traceWorker))
                worker = traceWorker;
            else if (!ExcelMcpService.IsSessionlessCommand(request.Command)
                && request.SessionId is { } sessionId && _sessions.TryGetValue(sessionId, out var owner))
                worker = owner;
            else if (request.Command == "session.close" && request.SessionId is { } closedId
                && _closedSessions.TryGetValue(closedId, out var end) && end == SessionEnd.Closed)
                return new ServiceResponse
                {
                    Success = true,
                    Command = request.Command,
                    SessionId = closedId,
                    Result = JsonSerializer.Serialize(new { success = true, sessionId = closedId, message = "Session already closed." },
                        ServiceProtocol.JsonOptions)
                };
            else if (!ExcelMcpService.IsSessionlessCommand(request.Command) && request.SessionId is { } lostId
                && _closedSessions.TryGetValue(lostId, out var lostEnd) && lostEnd == SessionEnd.WorkerExited)
                return CreateWorkerError(request,
                    new IOException("The workbook's Excel helper exited. Its last outcome could not be confirmed; no work was repeated."),
                    sent: true);
            else
            {
                lock (_utilityLock)
                {
                    if (_utility is null || _utility.IsFaulted || _utility.IsCanceled
                        || _utility.IsCompletedSuccessfully && (!_utility.Result.Registration.Worker.IsAlive()
                            || _utility.Result.Snapshot?.Releasing == true))
                        _utility = StartWorkerAsync(null, CancellationToken.None);
                }
                worker = await _utility.WaitAsync(request.CallerCancellation);
            }

            if (reservedPaths.Count > 0)
            {
                var associated = reservation with { Token = worker.Registration.PipeName };
                foreach (var path in reservedPaths)
                    _paths.TryUpdate(path, associated, reservation);
                reservation = associated;
            }
            request.CallerCancellation.ThrowIfCancellationRequested();
            if (_shutdown is not null)
                return ExcelWorkerHost.Failure(request, "SessionUnavailable", "The service is stopping. No new work was started.");
            if (request.Command == "workbook.save-as")
            {
                using var arguments = JsonDocument.Parse(request.Args ?? "{}");
                if (arguments.RootElement.TryGetProperty("targetPath", out var target)
                    && target.ValueKind == JsonValueKind.String && target.GetString() is { } targetPath)
                {
                    var normalized = FilePathValidation.NormalizeAbsoluteWindowsPath(targetPath);
                    if (!_paths.TryGetValue(normalized, out var existing) || existing.SessionId != request.SessionId)
                    {
                        if (!_paths.TryAdd(normalized, reservation))
                            return ExcelWorkerHost.Failure(request, "Conflict", "The Save As destination belongs to another session or pending request.");
                        reservedPaths.Add(normalized);
                    }
                }
            }
            if (!worker.Registration.Worker.IsAlive())
                throw new IOException("The Excel worker exited before this request was sent. Discover open workbooks to reconnect; no work was started.");
            sent = true;
            return await CompleteWorkerRequestAsync(request, worker, reservedPaths, reservation)
                .WaitAsync(request.CallerCancellation);
        }
        catch (Exception ex)
        {
            if (worker is not null && StartsWorkbookWorker(request.Command))
                _abandonedWorkers[worker.Registration.PipeName] = 0;
            return CreateWorkerError(request, ex, sent);
        }
        finally
        {
            if (!sent)
                foreach (var path in reservedPaths)
                    _paths.TryRemove(new KeyValuePair<string, PathReservation>(path, reservation));
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "The controller owns late worker completion after caller cancellation. It records exact results and confirmed state, and reports/logs transport failures rather than losing a faulted background task or retrying edits.")]
    private async Task<ServiceResponse> CompleteWorkerRequestAsync(
        ServiceRequest request, ExcelWorkerClient worker, List<string> reservedPaths, PathReservation reservation)
    {
        try
        {
            var execution = await worker.ExecuteWithAdmissionAsync(request, request.CallerCancellation);
            var response = execution.Response;
            if (!execution.Accepted)
                foreach (var path in reservedPaths)
                    _paths.TryRemove(new KeyValuePair<string, PathReservation>(path, reservation));
            // A status failure must not replace the exact workbook operation result.
            await RefreshWorkerAsync(worker, CancellationToken.None);
            lock (_shutdownLock)
            {
                if (!_workers.TryGetValue(worker.Registration.PipeName, out var current) || current != worker)
                    return response;
                if (response.SessionId is { } created && worker.Snapshot?.Sessions.Any(session => session.SessionId == created) == true)
                {
                    _sessions[created] = worker;
                    if (response.Success && request.Command == "session.attach"
                        && JsonNode.Parse(request.Args!)?["workbookId"]?.GetValue<string>() is { } selectionId
                        && _selections.TryGetValue(selectionId, out var selection) && selection.Worker == worker)
                        _selections[selectionId] = (worker, selection.Workbook with { SessionId = created });
                    var currentPaths = RecordedPaths(worker.Snapshot, created);
                    foreach (var path in reservedPaths)
                        if (currentPaths.Contains(path))
                            _paths[path] = new PathReservation("recorded", created);
                }
                if (response.Success && request.SessionId is { } reconciled
                    && worker.StatusError is null && worker.Snapshot is { StatusError: null } settledSnapshot
                    && !HasPendingWork(settledSnapshot)
                    && settledSnapshot.Sessions.FirstOrDefault(session => session.SessionId == reconciled)?.FilePath is not null)
                {
                    var currentPaths = RecordedPaths(settledSnapshot, reconciled);
                    foreach (var path in currentPaths)
                        _paths[path] = new PathReservation("recorded", reconciled);
                    foreach (var old in _paths.Where(path => path.Value.Token == "recorded" && path.Value.SessionId == reconciled
                        && !currentPaths.Contains(path.Key)))
                        _paths.TryRemove(old);
                }
                if (response.Success && request.Command == "session.close" && request.SessionId is { } closed)
                {
                    _closedSessions[closed] = SessionEnd.Closed;
                    _sessions.TryRemove(closed, out _);
                    foreach (var path in _paths.Where(path => path.Value.SessionId == closed))
                        _paths.TryRemove(path.Key, out _);
                }
                if (worker.StatusError is null && worker.Snapshot is { StatusError: null } snapshot
                    && !HasPendingWork(snapshot) && (response.Success || response.PartialState is null))
                {
                    var retainedPaths = snapshot.Sessions.SelectMany(session => RecordedPaths(snapshot, session.SessionId))
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);
                    foreach (var path in reservedPaths)
                        if (!retainedPaths.Contains(path))
                            _paths.TryRemove(new KeyValuePair<string, PathReservation>(path, reservation));
                }
            }
            if (request.Command is "session.close" or "session.test" or "sheet.copy-to-file" or "sheet.move-to-file"
                || !response.Success && reservedPaths.Count > 0)
                await ReleaseEmptyWorkerAsync(worker);
            return response;
        }
        catch (Exception ex)
        {
            if (StartsWorkbookWorker(request.Command))
                _abandonedWorkers[worker.Registration.PipeName] = 0;
            if (ex is OperationCanceledException && reservedPaths.Count > 0)
            {
                await RefreshWorkerAsync(worker, CancellationToken.None);
                lock (_shutdownLock)
                {
                    if (worker.StatusError is null && worker.Snapshot is { StatusError: null } snapshot && !HasPendingWork(snapshot))
                    {
                        var retainedPaths = snapshot.Sessions.SelectMany(session => RecordedPaths(snapshot, session.SessionId))
                            .ToHashSet(StringComparer.OrdinalIgnoreCase);
                        foreach (var path in reservedPaths)
                            if (!retainedPaths.Contains(path))
                                _paths.TryRemove(new KeyValuePair<string, PathReservation>(path, reservation));
                    }
                }
            }
            System.Diagnostics.Trace.TraceError("Excel worker request failed: {0}", ex);
            return CreateWorkerError(request, ex, sent: true);
        }
    }

    private static ServiceResponse CreateWorkerError(ServiceRequest request, Exception ex, bool sent) => new()
    {
        Success = false,
        Command = request.Command,
        SessionId = request.SessionId,
        ErrorCategory = ex is OperationCanceledException ? "Cancelled"
            : ex is TimeoutException ? "Timeout"
            : ex is ArgumentException or JsonException ? "InvalidInput" : "ServiceUnavailable",
        ExceptionType = ex.GetType().Name,
        ErrorMessage = $"{ex.GetType().Name}: {ex.Message}",
        PartialState = !sent || ex is ArgumentException or JsonException ? null : JsonSerializer.Serialize(new
        {
            outcome = "unknown",
            message = "The worker response could not be confirmed. Work may still be running; do not repeat edits automatically."
        }, ServiceProtocol.JsonOptions)
    };

    private static bool HasPendingWork(WorkerSnapshot snapshot) =>
        snapshot.PendingOperation is not null
        || JsonNode.Parse(snapshot.SessionList.Result!)!["sessions"]!.AsArray()
            .Any(session => session?["operationPending"]?.GetValue<bool>() == true);

    private static HashSet<string> RecordedPaths(WorkerSnapshot snapshot, string sessionId) =>
        snapshot.OwnershipKeys[sessionId].ToHashSet(StringComparer.OrdinalIgnoreCase);

    internal async Task<ExcelWorkerClient> StartWorkerAsync(ExcelProcessIdentity? excel, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ExcelWorkerClient worker;
        lock (_shutdownLock)
        {
            if (_shutdown is not null)
                throw new InvalidOperationException("The service is stopping. No Excel worker was started.");
            if (_workers.Count >= 32)
                throw new InvalidOperationException("The service has reached its Excel worker limit. Close unused sessions before retrying.");
            worker = ExcelWorkerClient.Launch(_launcher, _identity, _generation, excel);
            _workers[worker.Registration.PipeName] = worker;
            _newWorkers[worker.Registration.PipeName] = 0;
        }
        try
        {
            await _initializeWorker(worker);
        }
        catch
        {
            _abandonedWorkers[worker.Registration.PipeName] = 0;
            if (!worker.Registration.Worker.IsAlive())
                RetireExitedWorker(worker.Registration);
            throw;
        }
        return worker;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "An unused worker is released only after it reports no sessions and no pending work. Release failures retain its registration and appear as a status error, without changing the caller's workbook result.")]
    private async Task ReleaseEmptyWorkerAsync(ExcelWorkerClient worker)
    {
        if (worker.Snapshot is not { Sessions.Count: 0, PendingOperation: null, StatusError: null }
            || _sessions.Values.Contains(worker))
            return;
        try
        {
            var released = await worker.ReleaseAsync(CancellationToken.None);
            PublishProcessIdentity(worker.Registration);
            if (!released.Success)
                throw new InvalidOperationException(released.ErrorMessage);
            if (released.PartialState is not null)
                throw new InvalidOperationException($"The Excel worker did not confirm complete release: {released.PartialState}");
            _workers.TryRemove(worker.Registration.PipeName, out _);
            _newWorkers.TryRemove(worker.Registration.PipeName, out _);
            _abandonedWorkers.TryRemove(worker.Registration.PipeName, out _);
            foreach (var path in _paths.Where(entry => entry.Value.Token == worker.Registration.PipeName
                && entry.Value.SessionId is null))
                _paths.TryRemove(path);
            lock (_utilityLock)
                if (_utility is { IsCompletedSuccessfully: true } && _utility.Result == worker)
                    _utility = null;
            foreach (var selection in _selections.Where(entry => entry.Value.Worker == worker))
                _selections.TryRemove(selection.Key, out _);
        }
        catch (Exception ex)
        {
            worker.StatusError = $"The unused Excel worker could not be released: {ex.Message}";
            System.Diagnostics.Trace.TraceError("Could not release unused Excel worker: {0}", ex);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Recorded-status failures are exposed on session status and logged. They never replace a completed operation result or manufacture an empty healthy session list.")]
    private async Task RefreshWorkerAsync(ExcelWorkerClient worker, CancellationToken cancellationToken)
    {
        try
        {
            var snapshot = await worker.ReadSnapshotAsync(cancellationToken);
            PublishProcessIdentity(snapshot.Registration);
            lock (_shutdownLock)
            {
                if (!_workers.TryGetValue(worker.Registration.PipeName, out var current) || current != worker
                    || !ReferenceEquals(snapshot, worker.Snapshot))
                    return;
                foreach (var session in snapshot.Sessions)
                {
                    if (_closedSessions.ContainsKey(session.SessionId))
                        continue;
                    if (_newWorkers.ContainsKey(snapshot.Registration.PipeName))
                        _sessions[session.SessionId] = worker;
                    if (!string.IsNullOrWhiteSpace(session.FilePath))
                    {
                        var path = session.FilePath;
                        var currentPaths = RecordedPaths(snapshot, session.SessionId);
                        foreach (var ownedPath in currentPaths)
                            _paths[ownedPath] = new PathReservation("recorded", session.SessionId);
                        if (!HasPendingWork(snapshot) && snapshot.StatusError is null)
                        {
                            foreach (var selection in _selections.Where(entry => entry.Value.Worker == worker
                                && entry.Value.Workbook.SessionId == session.SessionId
                                && !currentPaths.Any(key => WorkbookLocation.MatchesReportedLocation(key,
                                    entry.Value.Workbook.Location))))
                                _selections[selection.Key] = (worker, selection.Value.Workbook with { Location = path });
                            foreach (var old in _paths.Where(entry => entry.Value.Token == "recorded"
                                && entry.Value.SessionId == session.SessionId
                                && !currentPaths.Contains(entry.Key)))
                                _paths.TryRemove(old);
                        }
                    }
                }
                if (!HasPendingWork(snapshot) && snapshot.StatusError is null)
                    foreach (var removed in _sessions.Where(entry => entry.Value == worker
                        && !snapshot.Sessions.Any(session => session.SessionId == entry.Key)))
                    {
                        _closedSessions[removed.Key] = SessionEnd.Closed;
                        _sessions.TryRemove(removed);
                        foreach (var path in _paths.Where(entry => entry.Value.SessionId == removed.Key))
                            _paths.TryRemove(path);
                    }
            }
        }
        catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
        {
            worker.StatusError = $"Excel worker status could not be refreshed: {ex.GetType().Name}: {ex.Message}";
            System.Diagnostics.Trace.TraceError("Excel worker status could not be refreshed: {0}", ex);
        }
    }

    private void PublishProcessIdentity(WorkerRegistration registration)
    {
        if (registration is { OwnsExcel: true, Excel: { } excel })
        {
            ExcelProcessObserved?.Invoke(excel);
            if (registration.ExcelPreserved)
                ExcelProcessPreserved?.Invoke(excel);
        }
    }

    private ServiceResponse SessionList()
    {
        var sessions = new JsonArray();
        foreach (var worker in _workers.Values)
        {
            if (worker.Snapshot is not { } snapshot)
                continue;
            var list = JsonNode.Parse(snapshot.SessionList.Result!)!.AsObject();
            foreach (var session in list["sessions"]!.AsArray())
            {
                if (session?["sessionId"]?.GetValue<string>() is not { } sessionId || !_sessions.ContainsKey(sessionId))
                    continue;
                var copy = session.DeepClone();
                if (snapshot.PendingOperation is not null)
                {
                    copy["operationPending"] = true;
                    copy["canClose"] = false;
                }
                if (worker.StatusError is not null || snapshot.Releasing)
                {
                    copy["canClose"] = false;
                    copy["blockingReason"] = "WorkerUnavailable";
                    copy["blockingMessage"] = worker.StatusError ?? "The Excel worker is being released.";
                }
                sessions.Add(copy);
            }
        }
        return new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new { success = true, sessions, count = sessions.Count }, ServiceProtocol.JsonOptions)
        };
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Discovery reports each failed application's exact inspection/connection error alongside other discovered workbooks. A failed application is never presented as an empty successful inspection.")]
    private async Task<ServiceResponse> DiscoverAsync(CancellationToken cancellationToken)
    {
        await _discovery.WaitAsync(cancellationToken);
        try
        {
            var native = RunningExcelWindows.Find();
            var retained = _registry.Read().Concat(_workers.Values.Select(worker => worker.Registration))
                .DistinctBy(record => record.PipeName).ToArray();
            var workbooks = new List<OpenWorkbookInfo>();
            var errors = new List<WorkbookDiscoveryError>(native.Errors);
            foreach (var window in native.Windows.DistinctBy(window => window.Process))
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var worker = _workers.Values.FirstOrDefault(worker =>
                        worker.Registration.Excel == window.Process && worker.Registration.Worker.IsAlive());
                    if (worker is null)
                    {
                        var record = retained.Where(record => record.Excel == window.Process)
                            .OrderByDescending(record => record.Worker.IsAlive()).FirstOrDefault();
                        if (record is not null)
                        {
                            if (!record.Worker.IsAlive())
                            {
                                errors.Add(new(window.Process.ProcessId,
                                    "The previous Excel worker exited without confirming its outcome. Its last work may have completed. This explicit discovery reconnects for inspection only; no edit was repeated."));
                                var exitedRecords = retained.Where(item =>
                                    item.Excel == window.Process && !item.Worker.IsAlive()).ToArray();
                                foreach (var exited in exitedRecords)
                                    RetireExitedWorker(exited, explicitDiscovery: true);
                                worker = await StartWorkerAsync(window.Process, cancellationToken);
                                foreach (var exited in exitedRecords)
                                    _registry.Remove(exited.PipeName);
                            }
                            else if (record.Build != WorkerRegistry.Build)
                            {
                                errors.Add(new(window.Process.ProcessId, "A retained Excel worker uses a different product build. It has not been replaced."));
                                continue;
                            }
                            else if (record.Controller != _identity && record.Controller.IsAlive())
                            {
                                errors.Add(new(window.Process.ProcessId, "This Excel application belongs to another running ExcelMcp controller."));
                                continue;
                            }
                            else
                            {
                                worker = new ExcelWorkerClient(record, _identity, _generation);
                                lock (_shutdownLock)
                                {
                                    if (_shutdown is not null)
                                        throw new InvalidOperationException("The service is stopping. No Excel application was claimed.");
                                    if (_workers.Count >= 32)
                                        throw new InvalidOperationException("The service has reached its Excel worker limit. Close unused sessions before retrying.");
                                    _workers[record.PipeName] = worker;
                                }
                                await worker.ClaimAsync(cancellationToken);
                            }
                        }
                        else
                            worker = await StartWorkerAsync(window.Process, cancellationToken);
                    }
                    if (worker.Snapshot?.Releasing == true)
                    {
                        errors.Add(new(window.Process.ProcessId, "The previous worker is still releasing Excel. No second Excel owner was started."));
                        continue;
                    }
                    if (worker.Snapshot is null)
                        await worker.ClaimAsync(cancellationToken);
                    var response = await CompleteWorkerRequestAsync(new ServiceRequest
                    {
                        Command = "session.discover",
                        CallerCancellation = cancellationToken
                    }, worker, [], new PathReservation("discovery")).WaitAsync(cancellationToken);
                    if (!response.Success)
                    {
                        errors.Add(new(window.Process.ProcessId, response.ErrorMessage ?? "Excel inspection failed."));
                        continue;
                    }
                    using var result = JsonDocument.Parse(response.Result!);
                    var discovered = result.RootElement.GetProperty("workbooks").Deserialize<List<OpenWorkbookInfo>>(ServiceProtocol.JsonOptions)!;
                    foreach (var workbook in discovered)
                    {
                        workbooks.Add(workbook);
                        _selections[workbook.WorkbookId] = (worker, workbook);
                    }
                    errors.AddRange(result.RootElement.GetProperty("errors").Deserialize<List<WorkbookDiscoveryError>>(ServiceProtocol.JsonOptions)!);
                }
                catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
                {
                    errors.Add(new(window.Process.ProcessId, $"{ex.GetType().Name}: {ex.Message}"));
                }
            }
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new { success = true, workbooks, errors, complete = errors.Count == 0 }, ServiceProtocol.JsonOptions)
            };
        }
        finally
        {
            _discovery.Release();
        }
    }

    private void RetireExitedWorker(WorkerRegistration registration, bool explicitDiscovery = false)
    {
        lock (_shutdownLock)
        {
            if (!_workers.TryGetValue(registration.PipeName, out var worker)
                || worker.Registration.Worker != registration.Worker || registration.Worker.IsAlive())
                return;
            if (!explicitDiscovery && registration.Excel is { } excel && !OwnedProcessGuard.TryConfirmExited(excel))
                return;
            if (!explicitDiscovery && registration.Excel is null
                && (worker.Snapshot?.Sessions.Count > 0 || _sessions.Values.Contains(worker)
                    || _paths.Values.Any(path => path.Token == registration.PipeName)))
                return;
            _workers.TryRemove(registration.PipeName, out _);
            _newWorkers.TryRemove(registration.PipeName, out _);
            _abandonedWorkers.TryRemove(registration.PipeName, out _);
            var sessions = _sessions.Where(entry => entry.Value == worker).Select(entry => entry.Key).ToHashSet(StringComparer.Ordinal);
            foreach (var session in sessions)
            {
                _closedSessions[session] = SessionEnd.WorkerExited;
                _sessions.TryRemove(session, out _);
            }
            foreach (var path in _paths.Where(entry => entry.Value.Token == registration.PipeName
                || entry.Value.SessionId is { } session && sessions.Contains(session)))
                _paths.TryRemove(path);
            foreach (var selection in _selections.Where(entry => entry.Value.Worker == worker))
                _selections.TryRemove(selection);
        }
    }

    internal Task<string?> ShutdownAsync()
    {
        lock (_shutdownLock)
            return _shutdown ??= Task.Run(StopWorkersAsync);
    }

    internal async Task<string?> RequestShutdownAsync()
    {
        var report = await ShutdownAsync();
        _host.RequestShutdown();
        _controlHost.RequestShutdown();
        return report;
    }

    internal void RevealSession(string sessionId)
    {
        if (_sessions.TryGetValue(sessionId, out var worker))
            _ = RevealSessionAsync(worker, sessionId);
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Caller-abandonment reveal is diagnosed separately without replacing the original cancellation or replaying a mutation.")]
    private static async Task RevealSessionAsync(ExcelWorkerClient worker, string sessionId)
    {
        try
        {
            var response = await worker.ExecuteAsync(new ServiceRequest
            {
                Command = "window.show",
                SessionId = sessionId
            }, CancellationToken.None);
            if (!response.Success)
                System.Diagnostics.Trace.TraceError("Could not reveal retained Excel work: {0}", response.ErrorMessage);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.TraceError("Could not reveal retained Excel work: {0}", ex);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Snapshot polling diagnoses unreachable workers without altering operation responses. Public status continues to use only previously recorded data.")]
    private async Task PollAsync()
    {
        while (!_stop.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(1), _stop.Token);
                await Task.WhenAll(_workers.Values.Select(async worker =>
                {
                    await RefreshWorkerAsync(worker, _stop.Token);
                    if (_abandonedWorkers.ContainsKey(worker.Registration.PipeName))
                    {
                        if (!worker.Registration.Worker.IsAlive())
                            RetireExitedWorker(worker.Registration);
                        else
                            await ReleaseEmptyWorkerAsync(worker);
                    }
                    else if (worker.Registration.Excel is { } excel
                        && OwnedProcessGuard.TryConfirmExited(excel))
                    {
                        if (!worker.Registration.Worker.IsAlive())
                        {
                            RetireExitedWorker(worker.Registration);
                            _registry.Remove(worker.Registration.PipeName);
                        }
                        else
                            await ReleaseEmptyWorkerAsync(worker);
                    }
                }));
            }
            catch (OperationCanceledException) when (_stop.IsCancellationRequested)
            {
                return;
            }
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Shutdown reports each worker that could not safely release Excel, retaining worker/process identities for explicit handover. It never waits indefinitely or kills Excel.")]
    private async Task<string?> StopWorkersAsync()
    {
        var reports = new ConcurrentBag<string>();
        await Task.WhenAll(_workers.Values.Select(async worker =>
        {
            try
            {
                var snapshot = await worker.ReadSnapshotAsync();
                if (HasPendingWork(snapshot))
                    throw new InvalidOperationException("An Excel operation is still pending. Its worker was retained for explicit reconnection.");
                var response = await worker.ReleaseAsync(CancellationToken.None);
                PublishProcessIdentity(worker.Registration);
                if (!response.Success)
                    throw new InvalidOperationException(response.ErrorMessage);
                if (response.PartialState is not null)
                    reports.Add(response.PartialState);
            }
            catch (Exception ex)
            {
                reports.Add($"Worker {worker.Registration.Worker.ProcessId}, Excel {worker.Registration.Excel?.ProcessId}: {ex.Message} " +
                    "Nothing was saved, discarded, or terminated to end this wait.");
            }
        }));
        return reports.IsEmpty ? null : string.Join(Environment.NewLine, reports);
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        ShutdownAsync().GetAwaiter().GetResult();
        _stop.Cancel();
        _poll.GetAwaiter().GetResult();
        _host.Dispose();
        _controlHost.Dispose();
        _stop.Dispose();
    }
}
