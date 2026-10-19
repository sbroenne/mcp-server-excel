using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Records every Excel process started while the scope is open so test cleanup
/// can stop the exact processes this test started if normal shutdown left them
/// running, for example behind an unanswered Excel dialog.
/// </summary>
internal sealed class TestExcelProcessScope : IDisposable
{
    private static readonly TimeSpan ExitWait = TimeSpan.FromSeconds(15);
    private static readonly TimeSpan TerminationWait = TimeSpan.FromSeconds(5);
    private readonly ConcurrentDictionary<ExcelProcessIdentity, byte> _identities = new();

    internal TestExcelProcessScope() =>
        SessionManager.ExcelProcessIdentityTracked += OnTracked;

    internal Exception? StopLeftovers() => StopLeftovers(_identities.Keys.ToArray());

    /// <summary>
    /// Waits for normal exit, then stops any remaining process by PID and start
    /// time. Returns a failure naming every leftover process, or null when all
    /// processes had already exited.
    /// </summary>
    internal static Exception? StopLeftovers(IReadOnlyCollection<ExcelProcessIdentity> identities)
    {
        SpinWait.SpinUntil(
            () => identities.All(OwnedProcessGuard.TryConfirmExited),
            ExitWait);

        var stopped = new List<string>();
        var notStopped = new List<string>();
        foreach (var identity in identities.Where(identity => !OwnedProcessGuard.TryConfirmExited(identity)))
        {
            if (OwnedProcessGuard.TryTerminate(identity, TimeSpan.Zero, TerminationWait, out _))
            {
                SessionManager.UntrackExcelProcess(identity);
                stopped.Add(identity.ToString());
            }
            else
            {
                notStopped.Add(identity.ToString());
            }
        }

        if (stopped.Count == 0 && notStopped.Count == 0)
        {
            return null;
        }

        return new InvalidOperationException(
            "Test cleanup found Excel still running after shutdown. " +
            $"Stopped: [{string.Join(", ", stopped)}]. Could not stop: [{string.Join(", ", notStopped)}].");
    }

    public void Dispose() =>
        SessionManager.ExcelProcessIdentityTracked -= OnTracked;

    private void OnTracked(ExcelProcessIdentity identity) => _identities.TryAdd(identity, 0);
}
