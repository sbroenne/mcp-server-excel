using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Microsoft.Win32.SafeHandles;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Tests.Shared;

internal sealed class TestRunExcelLifetime : IDisposable
{
    internal const string OwnershipDirectoryVariable = "EXCELMCP_TEST_OWNERSHIP_DIRECTORY";
    private static readonly object HostGate = new();
    private static bool _processExitRegistered;
    private readonly object _gate = new();
    private readonly string _journalPath;
    private readonly ExcelProcessIdentity _host;
    private readonly Dictionary<ExcelProcessIdentity, Process> _owned = [];
    private readonly SafeJobHandle _job;
    private readonly Func<ExcelProcessIdentity, bool> _identityExitProbe;
    private readonly Func<SafeProcessHandle, bool> _handleExitProbe;
    private bool _disposed;

    internal string JournalPath => _journalPath;
    internal static TestRunExcelLifetime? CurrentHost { get; private set; }

    internal static TestRunExcelLifetime GetOrCreateActive(TestRunExcelLifetime? current, string directory)
    {
        if (current is not null)
        {
            lock (current._gate)
            {
                if (!current._disposed) return current;
            }
        }
        return new TestRunExcelLifetime(directory);
    }

    internal static TestRunExcelLifetime StartForTestHost()
    {
        lock (HostGate)
        {
            var directory = Environment.GetEnvironmentVariable(OwnershipDirectoryVariable)
                ?? Path.Join(Path.GetTempPath(), "ExcelMcpTestOwnership", Guid.NewGuid().ToString("N"));
            CurrentHost = GetOrCreateActive(CurrentHost, directory);
            if (!_processExitRegistered)
            {
                AppDomain.CurrentDomain.ProcessExit += (_, _) => CurrentHost?.Dispose();
                _processExitRegistered = true;
            }
            return CurrentHost;
        }
    }

    internal TestRunExcelLifetime(
        string directory,
        Func<ExcelProcessIdentity, bool>? identityExitProbe = null,
        Func<SafeProcessHandle, bool>? handleExitProbe = null)
    {
        _identityExitProbe = identityExitProbe ?? OwnedProcessGuard.TryConfirmExited;
        _handleExitProbe = handleExitProbe ?? HasHandleExited;
        Directory.CreateDirectory(directory);
        using var host = Process.GetCurrentProcess();
        _host = new ExcelProcessIdentity(host.Id, host.StartTime.ToUniversalTime().ToFileTimeUtc());
        _journalPath = Path.Join(directory, $"host-{_host.ProcessId}-{_host.StartedAtUtcFileTime}-{Guid.NewGuid():N}.jsonl");
        Append(new OwnershipRecord("host", _host, null));
        _job = CreateOwnedJob();
        SessionManager.ExcelProcessIdentityTracked += OnTracked;
    }

    private void OnTracked(ExcelProcessIdentity identity)
    {
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_owned.ContainsKey(identity)) return;
            if (_identityExitProbe(identity))
            {
                Append(new OwnershipRecord("already-exited", _host, identity));
                return;
            }

            var process = Process.GetProcessById(identity.ProcessId);
            var retained = false;
            try
            {
                var handle = process.SafeHandle;
                if (!GetProcessTimes(handle, out var creationTime, out _, out _, out _))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
                if (creationTime != identity.StartedAtUtcFileTime)
                {
                    Append(new OwnershipRecord("identity-replaced", _host, identity));
                    return;
                }
                if (!string.Equals(Path.GetFileName(process.MainModule?.FileName),
                    "EXCEL.EXE", StringComparison.OrdinalIgnoreCase))
                {
                    // Managed ownership tests also register synthetic, non-Excel identities.
                    Append(new OwnershipRecord("non-excel-test-process", _host, identity));
                    return;
                }

                _owned.Add(identity, process);
                retained = true;
                if (!AssignProcessToJobObject(_job, handle))
                {
                    var error = Marshal.GetLastWin32Error();
                    Append(new OwnershipRecord("assignment-failed", _host, identity, error));
                    throw new Win32Exception(error,
                        "Could not protect the recorded Excel process against testhost termination.");
                }
                Append(new OwnershipRecord("excel", _host, identity));
            }
            finally
            {
                if (!retained) process.Dispose();
            }
        }
    }

    private void Append(OwnershipRecord record)
    {
        var bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(record) + "\n");
        using var stream = new FileStream(
            _journalPath, FileMode.Append, FileAccess.Write, FileShare.Read,
            4096, FileOptions.WriteThrough);
        stream.Write(bytes);
        stream.Flush(flushToDisk: true);
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _disposed = true;
            SessionManager.ExcelProcessIdentityTracked -= OnTracked;
            ExcelProcessIdentity[] remaining;
            try
            {
                remaining = _owned
                    .Where(owned => !HasExited(owned.Key, owned.Value.SafeHandle))
                    .Select(owned => owned.Key)
                    .ToArray();
                foreach (var identity in remaining)
                {
                    Append(new OwnershipRecord("normal-teardown-failed", _host, identity));
                }
                Append(new OwnershipRecord("disposed", _host, null));
            }
            finally
            {
                _job.Dispose();
                foreach (var process in _owned.Values) process.Dispose();
                _owned.Clear();
            }

            if (remaining.Length > 0)
            {
                throw new InvalidOperationException(
                    $"Owned Excel survived normal testhost teardown before the job fallback: {string.Join(", ", remaining)}.");
            }
        }
    }

    internal bool HasExited(ExcelProcessIdentity identity)
    {
        lock (_gate)
        {
            return HasExited(identity, _owned[identity].SafeHandle);
        }
    }

    internal bool HasExited(ExcelProcessIdentity identity, SafeProcessHandle originalHandle)
    {
        try
        {
            return _handleExitProbe(originalHandle);
        }
        catch (Win32Exception ex)
        {
            throw new InvalidOperationException($"Cannot determine exit status for captured Excel identity {identity}.", ex);
        }
    }

    private static bool HasHandleExited(SafeProcessHandle handle) =>
        WaitForSingleObject(handle, 0) switch
        {
            0 => true,
            0x102 => false,
            _ => throw new Win32Exception(Marshal.GetLastWin32Error())
        };

    private static SafeJobHandle CreateOwnedJob()
    {
        var job = CreateJobObject(IntPtr.Zero, null);
        if (job.IsInvalid)
        {
            var error = Marshal.GetLastWin32Error();
            job.Dispose();
            throw new Win32Exception(error);
        }

        var limits = new JobLimits();
        limits.Basic.LimitFlags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        if (!SetInformationJobObject(job, 9, ref limits, (uint)Marshal.SizeOf<JobLimits>()))
        {
            var error = Marshal.GetLastWin32Error();
            job.Dispose();
            throw new Win32Exception(error);
        }
        return job;
    }

    private sealed class SafeJobHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public SafeJobHandle() : base(ownsHandle: true) { }
        protected override bool ReleaseHandle() => CloseHandle(handle);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct BasicLimits
    {
        internal long ProcessTime, JobTime;
        internal uint LimitFlags;
        internal nuint MinWorkingSet, MaxWorkingSet;
        internal uint ActiveProcessLimit;
        internal nuint Affinity;
        internal uint PriorityClass, SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IoCounters
    {
        internal ulong ReadOperations, WriteOperations, OtherOperations;
        internal ulong ReadBytes, WriteBytes, OtherBytes;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JobLimits
    {
        internal BasicLimits Basic;
        internal IoCounters Io;
        internal nuint ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory;
    }

    [DllImport("kernel32.dll", EntryPoint = "CreateJobObjectW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeJobHandle CreateJobObject(IntPtr attributes, string? name);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetInformationJobObject(SafeJobHandle job, int informationClass, ref JobLimits limits, uint length);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AssignProcessToJobObject(SafeJobHandle job, SafeProcessHandle process);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetProcessTimes(SafeProcessHandle process, out long creation, out long exit, out long kernel, out long user);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint WaitForSingleObject(SafeProcessHandle handle, uint milliseconds);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);

    internal sealed record OwnershipRecord(
        string Kind,
        ExcelProcessIdentity Host,
        ExcelProcessIdentity? Excel,
        int? ErrorCode = null);
}
