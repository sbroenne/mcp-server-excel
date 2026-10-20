using System.Collections.Concurrent;
using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Tests.Shared;

/// <summary>
/// Records every Excel process started during a test run and fails the run when any
/// of them is still running at the end. It never closes or stops Excel, and it marks
/// reported processes so the product's own exit handling leaves them running too.
/// </summary>
internal sealed class TestRunExcelLeftoverCheck : IDisposable
{
    private static readonly TimeSpan ExitWait = TimeSpan.FromSeconds(15);
    private static readonly object Gate = new();
    private static TestRunExcelLeftoverCheck? _current;
    private readonly ConcurrentDictionary<ExcelProcessIdentity, Process> _started = new();
    private readonly Func<Process, bool> _isExcel;

    internal TestRunExcelLeftoverCheck(Func<Process, bool>? isExcel = null) =>
        _isExcel = isExcel ?? (process => string.Equals(process.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase));

    internal static void Start()
    {
        lock (Gate)
        {
            if (_current is null)
            {
                _current = new TestRunExcelLeftoverCheck();
                SessionManager.ExcelProcessIdentityTracked += _current.Record;
            }
        }
    }

    internal static void Finish()
    {
        TestRunExcelLeftoverCheck check;
        lock (Gate)
        {
            check = _current ?? throw new InvalidOperationException("The Excel leftover check was not started.");
            _current = null;
        }

        SessionManager.ExcelProcessIdentityTracked -= check.Record;
        using (check)
        {
            var remaining = check.FindRemaining(ExitWait);
            if (remaining.Length > 0)
            {
                throw new InvalidOperationException(
                    $"Excel started by this test run is still running: {string.Join(", ", remaining)}. " +
                    "Tests do not close or stop Excel; find the test that left it open.");
            }
        }
    }

    internal void Record(ExcelProcessIdentity identity)
    {
        // Tracking tests register the test host itself as a stand-in for Excel.
        if (identity.ProcessId == Environment.ProcessId)
        {
            return;
        }

        Process process;
        try
        {
            process = Process.GetProcessById(identity.ProcessId);
        }
        catch (ArgumentException)
        {
            return; // Already exited.
        }

        var retained = false;
        try
        {
            // Keep the original handle so a later process reusing this ID cannot be mistaken for it.
            _ = process.SafeHandle;
            retained = process.StartTime.ToUniversalTime().ToFileTimeUtc() == identity.StartedAtUtcFileTime
                && _isExcel(process)
                && _started.TryAdd(identity, process);
        }
        catch (InvalidOperationException) when (process.HasExited)
        {
        }
        finally
        {
            if (!retained)
            {
                process.Dispose();
            }
        }
    }

    internal ExcelProcessIdentity[] FindRemaining(TimeSpan wait)
    {
        var elapsed = Stopwatch.StartNew();
        var remaining = _started
            .Where(started => !started.Value.WaitForExit(
                wait > elapsed.Elapsed ? wait - elapsed.Elapsed : TimeSpan.Zero))
            .Select(started => started.Key)
            .ToArray();
        foreach (var identity in remaining)
        {
            SessionManager.PreserveExcelProcess(identity);
        }

        return remaining;
    }

    public void Dispose()
    {
        foreach (var process in _started.Values)
        {
            process.Dispose();
        }

        _started.Clear();
    }
}
