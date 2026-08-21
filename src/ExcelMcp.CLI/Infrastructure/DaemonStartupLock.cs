namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Serializes daemon startup and owned teardown for one CLI pipe.
/// </summary>
internal static class DaemonStartupLock
{
    internal static readonly TimeSpan Timeout = TimeSpan.FromSeconds(11);

    internal static string GetDaemonMutexName(string pipeName) =>
        $"ExcelMcpCli_Daemon_{DaemonPipeIdentity.GetHash(pipeName)}";

    internal static string GetLegacyDaemonMutexName(string pipeName) =>
        $"ExcelMcpCli_{pipeName}";

    internal static IReadOnlyList<string> GetLegacyDaemonMutexNames(string pipeName) =>
        DaemonPipeIdentity.GetLegacyCaseVariants(pipeName)
            .Select(GetLegacyDaemonMutexName)
            .ToList();

    internal static string GetStartupMutexName(string pipeName) =>
        $"ExcelMcpCli_Startup_{DaemonPipeIdentity.GetHash(pipeName)}";

    internal static string GetStartingMarkerName(string pipeName) =>
        $"ExcelMcpCli_Starting_{DaemonPipeIdentity.GetHash(pipeName)}";

    internal static Task<T> WithLockAsync<T>(
        string pipeName,
        Func<Task<T>> action,
        CancellationToken cancellationToken)
    {
        return Task.Run(() =>
        {
            using var startupMutex = new Mutex(
                initiallyOwned: false,
                GetStartupMutexName(pipeName),
                out _);
            var startupLockAcquired = false;
            try
            {
                try
                {
                    startupLockAcquired = startupMutex.WaitOne(Timeout);
                }
                catch (AbandonedMutexException)
                {
                    startupLockAcquired = true;
                }

                if (!startupLockAcquired)
                {
                    throw new TimeoutException(
                        $"Could not acquire the CLI startup lock for pipe '{pipeName}' " +
                        $"within {Timeout.TotalSeconds:0} seconds.");
                }

                return action().GetAwaiter().GetResult();
            }
            finally
            {
                if (startupLockAcquired)
                {
                    startupMutex.ReleaseMutex();
                }
            }
        }, cancellationToken);
    }
}
