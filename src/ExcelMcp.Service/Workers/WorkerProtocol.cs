using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service.Workers;

internal sealed record WorkerProcess(int ProcessId, long StartedAtUtcFileTime)
{
    internal static Action<Process>? BeforeIdentityReadForTests { get; set; }

    internal static WorkerProcess Current
    {
        get
        {
            using var process = Process.GetCurrentProcess();
            return From(process);
        }
    }

    internal static WorkerProcess From(Process process) =>
        new(process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());

    internal bool IsAlive()
    {
        Process process;
        try
        {
            process = Process.GetProcessById(ProcessId);
        }
        catch (ArgumentException)
        {
            return false;
        }
        using (process)
        {
            try
            {
                _ = process.Handle;
                if (process.HasExited)
                    return false;
                BeforeIdentityReadForTests?.Invoke(process);
                using var current = Process.GetCurrentProcess();
                return process.SessionId == current.SessionId
                    && process.StartTime.ToUniversalTime().ToFileTimeUtc() == StartedAtUtcFileTime;
            }
            catch (InvalidOperationException) when (process.HasExited)
            {
                return false;
            }
        }
    }
}

internal sealed record WorkerRegistration(
    string PipeName,
    WorkerProcess Worker,
    WorkerProcess Controller,
    string Generation,
    string Build,
    ExcelProcessIdentity? Excel,
    bool OwnsExcel = false,
    bool ExcelPreserved = false);

internal sealed record WorkerClaim(WorkerProcess Controller, string Generation, string Build);
internal sealed record WorkerCall(string OperationId, string Generation, ServiceRequest Request);
internal sealed record WorkerExecution(ServiceResponse Response, bool Accepted);
internal sealed record WorkerRelease(ServiceResponse Response, WorkerSnapshot Snapshot);
internal sealed record WorkerSnapshot(
    WorkerRegistration Registration,
    IReadOnlyList<SessionDescriptor> Sessions,
    ServiceResponse SessionList,
    string? PendingOperation,
    string? PendingCommand,
    IReadOnlyDictionary<string, IReadOnlyList<string>> OwnershipKeys,
    string? StatusError = null,
    bool Releasing = false,
    long Revision = 0);

[StreamJsonRpc.JsonRpcContract]
[PolyType.GenerateShape(IncludeMethods = PolyType.MethodShapeFlags.PublicInstance)]
internal partial interface IExcelWorkerRpc
{
    Task<WorkerSnapshot> ClaimAsync(WorkerClaim claim, CancellationToken cancellationToken = default);
    Task<WorkerSnapshot> SnapshotAsync(string generation, CancellationToken cancellationToken = default);
    Task<WorkerExecution> ExecuteAsync(WorkerCall call, CancellationToken cancellationToken = default);
    Task<WorkerRelease> ReleaseAsync(string generation, CancellationToken cancellationToken = default);
}

internal static class WorkerPipeIdentity
{
    internal static void VerifyServer(SafePipeHandle handle, WorkerProcess expected)
    {
        if (!GetNamedPipeServerProcessId(handle, out var processId))
            throw new Win32Exception(Marshal.GetLastWin32Error());
        if (processId != expected.ProcessId || !expected.IsAlive())
            throw new InvalidOperationException("The Excel worker pipe belongs to a different process.");
    }

    internal static int GetClient(SafePipeHandle handle)
    {
        if (!GetNamedPipeClientProcessId(handle, out var processId))
            throw new Win32Exception(Marshal.GetLastWin32Error());
        return checked((int)processId);
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetNamedPipeServerProcessId(SafePipeHandle pipe, out uint processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetNamedPipeClientProcessId(SafePipeHandle pipe, out uint processId);
}
