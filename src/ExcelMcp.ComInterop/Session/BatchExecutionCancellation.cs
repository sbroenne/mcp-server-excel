namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Carries transport cancellation into Core commands without exposing transport
/// concerns on every generated command contract.
/// </summary>
public static class BatchExecutionCancellation
{
    private static readonly AsyncLocal<CancellationToken> CurrentToken = new();
    private static readonly AsyncLocal<bool> CooperativeCleanup = new();

    /// <summary>
    /// Gets the cancellation token for the current service request.
    /// </summary>
    public static CancellationToken Current => CurrentToken.Value;

    /// <summary>
    /// Gets whether the current operation must finish bounded cleanup before returning.
    /// </summary>
    public static bool RequiresCooperativeCleanup => CooperativeCleanup.Value;

    /// <summary>
    /// Pushes a cancellation token for the current asynchronous request flow.
    /// </summary>
    public static IDisposable Push(
        CancellationToken cancellationToken,
        bool requiresCooperativeCleanup = false)
    {
        CancellationToken previous = CurrentToken.Value;
        bool previousCleanup = CooperativeCleanup.Value;
        CurrentToken.Value = cancellationToken;
        CooperativeCleanup.Value = requiresCooperativeCleanup;
        return new CancellationScope(previous, previousCleanup);
    }

    private sealed class CancellationScope(
        CancellationToken previous,
        bool previousCleanup) : IDisposable
    {
        private bool _disposed;

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            CurrentToken.Value = previous;
            CooperativeCleanup.Value = previousCleanup;
            _disposed = true;
        }
    }
}
