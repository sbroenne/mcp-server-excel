using System.Text.Json;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class Program
{
    private static async Task<int> Main()
    {
        var pipeName = Environment.GetEnvironmentVariable("EXCELMCP_CLI_PIPE")
            ?? ServiceSecurity.GetCliPipeName();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(45));

        try
        {
            var result = await DaemonStartupLock.WithLockAsync(
                pipeName,
                () => PreBuildProcessCleanup.CleanupWithGracefulShutdownAsync(
                    pipeName,
                    timeout.Token),
                timeout.Token);
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                success = result.Success,
                daemonMatched = result.DaemonMatched,
                error = result.ErrorMessage
            }));
            return result.Success ? 0 : 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                success = false,
                error = ex.Message
            }));
            return 1;
        }
    }
}
