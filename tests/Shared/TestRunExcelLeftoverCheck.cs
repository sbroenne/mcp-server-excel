using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Tests.Shared;

/// <summary>
/// Records every Excel process started during a test run and fails the run when any
/// of them is still running at the end. It never closes or stops Excel.
/// </summary>
internal sealed class TestRunExcelLeftoverCheck
{
    private static readonly TimeSpan ExitWait = TimeSpan.FromSeconds(15);
    private static readonly object Gate = new();
    private static TestRunExcelLeftoverCheck? _current;
    private readonly ConcurrentDictionary<ExcelProcessIdentity, byte> _started = new();

    private TestRunExcelLeftoverCheck() =>
        SessionManager.ExcelProcessIdentityTracked += OnTracked;

    internal static void Start()
    {
        lock (Gate)
        {
            _current ??= new TestRunExcelLeftoverCheck();
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

        SessionManager.ExcelProcessIdentityTracked -= check.OnTracked;
        var started = check._started.Keys.ToArray();
        SpinWait.SpinUntil(() => started.All(OwnedProcessGuard.TryConfirmExited), ExitWait);
        var remaining = started.Where(identity => !OwnedProcessGuard.TryConfirmExited(identity)).ToArray();
        if (remaining.Length > 0)
        {
            throw new InvalidOperationException(
                $"Excel started by this test run is still running: {string.Join(", ", remaining)}. " +
                "Tests do not close or stop Excel; find the test that left it open.");
        }
    }

    private void OnTracked(ExcelProcessIdentity identity)
    {
        // Tracking tests register the test host itself as a stand-in for Excel.
        if (identity.ProcessId != Environment.ProcessId)
        {
            _started.TryAdd(identity, 0);
        }
    }
}
