using System.Globalization;
using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using StreamJsonRpc;

namespace Sbroenne.ExcelMcp.Service.Workers;

internal sealed class ExcelWorkerClient
{
    internal static readonly TimeSpan ControlTimeout = TimeSpan.FromSeconds(2);
    internal static readonly TimeSpan StartupTimeout = TimeSpan.FromSeconds(15);
    private readonly WorkerProcess _controller;
    private readonly string _generation;
    private readonly Lock _snapshotLock = new();
    internal WorkerRegistration Registration { get; private set; }
    internal WorkerSnapshot? Snapshot { get; private set; }
    internal string? StatusError { get; set; }

    internal ExcelWorkerClient(WorkerRegistration registration, WorkerProcess controller, string generation)
    {
        Registration = registration;
        _controller = controller;
        _generation = generation;
    }

    internal static ExcelWorkerClient Launch(
        ProductProcessLauncher launcher, WorkerProcess controller, string generation, ExcelProcessIdentity? excel)
    {
        var pipe = $"excelmcp-worker-{Guid.NewGuid():N}";
        using var process = launcher.Start(
        [
            ExcelWorkerHost.Mode, pipe,
            controller.ProcessId.ToString(CultureInfo.InvariantCulture),
            controller.StartedAtUtcFileTime.ToString(CultureInfo.InvariantCulture),
            generation, WorkerRegistry.Build,
            (excel?.ProcessId ?? 0).ToString(CultureInfo.InvariantCulture),
            (excel?.StartedAtUtcFileTime ?? 0).ToString(CultureInfo.InvariantCulture)
        ]);
        return new ExcelWorkerClient(
            new(pipe, WorkerProcess.From(process), controller, generation, WorkerRegistry.Build, excel),
            controller, generation);
    }

    internal async Task InitializeAsync() => RecordSnapshot(await CallAsync(
        rpc => rpc.ClaimAsync(new(_controller, _generation, WorkerRegistry.Build), CancellationToken.None),
        CancellationToken.None, StartupTimeout));

    internal async Task<WorkerSnapshot> ClaimAsync(CancellationToken cancellationToken) =>
        RecordSnapshot(await CallAsync(
            rpc => rpc.ClaimAsync(new(_controller, _generation, WorkerRegistry.Build), cancellationToken),
            cancellationToken, ControlTimeout));

    internal async Task<WorkerSnapshot> ReadSnapshotAsync(CancellationToken cancellationToken = default) =>
        RecordSnapshot(await CallAsync(rpc => rpc.SnapshotAsync(_generation, cancellationToken),
            cancellationToken, ControlTimeout));

    internal WorkerSnapshot RecordSnapshot(WorkerSnapshot snapshot)
    {
        if (snapshot.Registration.PipeName != Registration.PipeName
            || snapshot.Registration.Worker != Registration.Worker
            || snapshot.Registration.Controller != _controller
            || snapshot.Registration.Generation != _generation
            || snapshot.Registration.Build != WorkerRegistry.Build)
            throw new InvalidOperationException("The Excel worker returned a mismatched process identity or build.");
        lock (_snapshotLock)
        {
            if (Snapshot is { } current && current.Revision >= snapshot.Revision)
                return current;
            Registration = snapshot.Registration;
            Snapshot = snapshot;
            StatusError = snapshot.StatusError;
            return snapshot;
        }
    }

    internal async Task<ServiceResponse> ExecuteAsync(ServiceRequest request, CancellationToken cancellationToken) =>
        (await ExecuteWithAdmissionAsync(request, cancellationToken)).Response;

    internal async Task<WorkerExecution> ExecuteWithAdmissionAsync(ServiceRequest request, CancellationToken cancellationToken)
    {
        var execution = await CallAsync(rpc => rpc.ExecuteAsync(new(Guid.NewGuid().ToString("N"), _generation, request), cancellationToken),
            CancellationToken.None, ServiceClient.DefaultRequestTimeout);
        return execution with { Snapshot = RecordSnapshot(execution.Snapshot) };
    }

    internal async Task<ServiceResponse> ReleaseAsync(CancellationToken cancellationToken)
    {
        var released = await CallAsync(rpc => rpc.ReleaseAsync(_generation, cancellationToken),
            cancellationToken, TimeSpan.FromSeconds(5));
        RecordSnapshot(released.Snapshot);
        return released.Response;
    }

    private Process OpenWorkerProcess()
    {
        try
        {
            return Process.GetProcessById(Registration.Worker.ProcessId);
        }
        catch (ArgumentException ex)
        {
            throw new IOException($"Excel worker {Registration.Worker.ProcessId} has exited. Its last workbook outcome is unconfirmed; no work was repeated.", ex);
        }
    }

    private async Task<T> CallAsync<T>(
        Func<IExcelWorkerRpc, Task<T>> invoke, CancellationToken cancellationToken, TimeSpan timeout)
    {
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(timeout);
        using var process = OpenWorkerProcess();
        if (WorkerProcess.From(process) != Registration.Worker)
            throw new InvalidOperationException("The recorded Excel worker process has been replaced.");
        using var monitor = CancellationTokenSource.CreateLinkedTokenSource(deadline.Token);
        var exited = process.WaitForExitAsync(monitor.Token);
        try
        {
            await using var pipe = ServiceSecurity.CreateClient(Registration.PipeName);
            var connecting = pipe.ConnectAsync(deadline.Token);
            if (await Task.WhenAny(connecting, exited) == exited && exited.IsCompletedSuccessfully)
                throw new IOException($"Excel worker {process.Id} exited before connecting (exit code {process.ExitCode}). Ensure the controller and worker were built together.");
            await connecting;
            WorkerPipeIdentity.VerifyServer(pipe.SafePipeHandle, Registration.Worker);
            var proxy = JsonRpc.Attach<IExcelWorkerRpc>(pipe);
            using var connection = (IDisposable)proxy;
            var response = invoke(proxy);
            var completed = await Task.WhenAny(response, exited).WaitAsync(deadline.Token);
            if (completed == exited && exited.IsCompletedSuccessfully && !response.IsCompleted)
                throw new IOException($"Excel worker {process.Id} exited while its response was pending (exit code {process.ExitCode}). Its workbook outcome is unconfirmed; no work was repeated.");
            return await response.WaitAsync(deadline.Token);
        }
        catch (OperationCanceledException ex) when (!cancellationToken.IsCancellationRequested)
        {
            throw new TimeoutException("The Excel worker did not respond within the allotted wait. Its work was not repeated or terminated.", ex);
        }
        finally
        {
            await monitor.CancelAsync();
        }
    }
}
