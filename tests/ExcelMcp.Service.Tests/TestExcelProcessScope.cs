using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Records every Excel process started while the scope is open so a test can fail
/// when normal shutdown left one running. It never closes or stops Excel.
/// </summary>
internal sealed class TestExcelProcessScope : IDisposable
{
    private static readonly TimeSpan ExitWait = TimeSpan.FromSeconds(15);
    private readonly ConcurrentDictionary<ExcelProcessIdentity, byte> _identities = new();

    internal TestExcelProcessScope() =>
        SessionManager.ExcelProcessIdentityTracked += OnTracked;

    internal Exception? FindLeftovers() => FindLeftovers(_identities.Keys.ToArray());

    /// <summary>
    /// Waits for normal exit. Returns a failure naming every process still running,
    /// or null when all of them exited.
    /// </summary>
    internal static Exception? FindLeftovers(IReadOnlyCollection<ExcelProcessIdentity> identities)
    {
        SpinWait.SpinUntil(
            () => identities.All(OwnedProcessGuard.TryConfirmExited),
            ExitWait);

        var remaining = identities.Where(identity => !OwnedProcessGuard.TryConfirmExited(identity)).ToArray();
        return remaining.Length == 0
            ? null
            : new InvalidOperationException(
                $"Excel is still running after shutdown: [{string.Join(", ", remaining)}]. Tests do not stop it.");
    }

    public void Dispose() =>
        SessionManager.ExcelProcessIdentityTracked -= OnTracked;

    private void OnTracked(ExcelProcessIdentity identity) => _identities.TryAdd(identity, 0);
}
