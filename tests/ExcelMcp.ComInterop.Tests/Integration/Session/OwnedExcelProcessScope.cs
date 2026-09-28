using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

internal sealed class OwnedExcelProcessScope : IDisposable
{
    private static readonly AsyncLocal<OwnedExcelProcessScope?> Current = new();
    private readonly OwnedExcelProcessScope? _previous = Current.Value;
    private readonly ConcurrentDictionary<ExcelProcessIdentity, byte> _owned = new();

    internal OwnedExcelProcessScope()
    {
        Current.Value = this;
        SessionManager.ExcelProcessIdentityTracked += OnTracked;
    }

    private void OnTracked(ExcelProcessIdentity identity)
    {
        if (ReferenceEquals(Current.Value, this))
        {
            _owned.TryAdd(identity, 0);
        }
    }

    internal void AssertAllExited(bool expectProcess = true)
    {
        if (expectProcess)
        {
            Assert.NotEmpty(_owned);
        }

        var exited = SpinWait.SpinUntil(
            () => _owned.Keys.All(OwnedProcessGuard.TryConfirmExited),
            TimeSpan.FromSeconds(15));
        var remaining = _owned.Keys.Where(identity => !OwnedProcessGuard.TryConfirmExited(identity)).ToArray();
        foreach (var identity in remaining)
        {
            OwnedProcessGuard.TryTerminate(identity, TimeSpan.Zero, TimeSpan.FromSeconds(5), out _);
        }

        Assert.True(exited, $"Owned Excel processes survived shutdown: {string.Join(", ", remaining)}");
    }

    public void Dispose()
    {
        SessionManager.ExcelProcessIdentityTracked -= OnTracked;
        Current.Value = _previous;
    }
}
