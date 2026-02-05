using System.Diagnostics;

namespace Sbroenne.ExcelMcp.ComInterop.ServiceClient;

/// <summary>
/// Launches and manages the ExcelMCP Service process.
/// Used by both CLI and MCP Server to ensure the service is running.
/// </summary>
public static class ServiceLauncher
{
    private static readonly SemaphoreSlim _startLock = new(1, 1);

    /// <summary>
    /// Environment variable to override the service executable path.
    /// Useful for development and testing.
    /// </summary>
    public const string ServicePathEnvVar = "EXCELMCP_SERVICE_PATH";

    /// <summary>
    /// Ensures the service is running, starting it if necessary.
    /// Thread-safe - multiple concurrent calls will only start one instance.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>True if service is running (or was started), false if failed to start</returns>
    public static async Task<bool> EnsureServiceRunningAsync(CancellationToken cancellationToken = default)
    {
        // Quick check without lock
        if (await IsServiceAliveAsync(cancellationToken))
        {
            return true;
        }

        // Acquire lock for starting
        await _startLock.WaitAsync(cancellationToken);
        try
        {
            // Double-check after acquiring lock
            if (await IsServiceAliveAsync(cancellationToken))
            {
                return true;
            }

            return await StartServiceAsync(cancellationToken);
        }
        finally
        {
            _startLock.Release();
        }
    }

    /// <summary>
    /// Checks if the service is running and responsive.
    /// </summary>
    public static async Task<bool> IsServiceAliveAsync(CancellationToken cancellationToken = default)
    {
        // First check lock file
        if (!ServiceSecurity.IsServiceProcessRunning())
        {
            return false;
        }

        // Then ping
        using var client = new ExcelServiceClient("launcher", connectTimeout: TimeSpan.FromSeconds(2));
        return await client.PingAsync(cancellationToken);
    }

    /// <summary>
    /// Starts the service as a background process.
    /// </summary>
    public static async Task<bool> StartServiceAsync(CancellationToken cancellationToken = default)
    {
        var startInfo = GetServiceStartInfo();
        if (startInfo == null)
        {
            return false;
        }

        try
        {
            var process = Process.Start(startInfo);
            if (process == null)
            {
                return false;
            }

            // Wait for service to start
            await Task.Delay(500, cancellationToken);

            // Verify it's running with retries
            for (int i = 0; i < 10; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                using var client = new ExcelServiceClient("launcher", connectTimeout: TimeSpan.FromSeconds(1));
                if (await client.PingAsync(cancellationToken))
                {
                    return true;
                }

                await Task.Delay(200, cancellationToken);
            }

            return false;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            return false;
        }
    }

    private static ProcessStartInfo? GetServiceStartInfo()
    {
        // Check for override path (development/testing)
        var overridePath = Environment.GetEnvironmentVariable(ServicePathEnvVar);
        if (!string.IsNullOrEmpty(overridePath) && File.Exists(overridePath))
        {
            return new ProcessStartInfo
            {
                FileName = overridePath,
                Arguments = "service run",
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
        }

        // Try to find excelcli.exe - prioritize same directory (unified package)
        var searchPaths = new[]
        {
            // Same directory as current process (unified package - excelcli bundled with mcp-excel)
            Path.Combine(AppContext.BaseDirectory, "excelcli.exe"),
            // NuGet tools location (global tool install)
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".dotnet", "tools", "excelcli.exe"),
            // Parent directory (legacy/alternative layouts)
            Path.Combine(AppContext.BaseDirectory, "..", "excelcli.exe"),
            // Program Files
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ExcelMCP", "excelcli.exe"),
            // LocalAppData
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelMCP", "excelcli.exe")
        };

        foreach (var path in searchPaths)
        {
            var normalizedPath = Path.GetFullPath(path);
            if (File.Exists(normalizedPath))
            {
                return new ProcessStartInfo
                {
                    FileName = normalizedPath,
                    Arguments = "service run",
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
            }
        }

        // Development mode: try 'dotnet run' from CLI project
        var entryAssembly = System.Reflection.Assembly.GetEntryAssembly();
        if (entryAssembly != null)
        {
            // Try to find CLI project relative to current location
            var baseDir = AppContext.BaseDirectory;
            var possibleCliDirs = new[]
            {
                // MCP Server is sibling to CLI
                Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "ExcelMcp.CLI")),
                Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "ExcelMcp.CLI")),
                Path.GetFullPath(Path.Combine(baseDir, "..", "..", "ExcelMcp.CLI")),
            };

            foreach (var cliDir in possibleCliDirs)
            {
                var csprojPath = Path.Combine(cliDir, "ExcelMcp.CLI.csproj");
                if (File.Exists(csprojPath))
                {
                    return new ProcessStartInfo
                    {
                        FileName = "dotnet",
                        Arguments = $"run --project \"{csprojPath}\" -- service run",
                        UseShellExecute = true,
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    };
                }
            }
        }

        return null;
    }
}


