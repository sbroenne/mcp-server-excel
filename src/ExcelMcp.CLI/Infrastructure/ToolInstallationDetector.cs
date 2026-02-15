using System.Diagnostics;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// Detects whether the CLI is installed as a global or local .NET tool.
/// </summary>
internal static class ToolInstallationDetector
{
    /// <summary>
    /// Gets the installation type (Global or Local).
    /// </summary>
    public static InstallationType GetInstallationType()
    {
        try
        {
            // Get the current executable path
            var executablePath = Environment.ProcessPath;
            if (string.IsNullOrEmpty(executablePath))
                return InstallationType.Unknown;

            // Global tools are installed in:
            // Windows: %USERPROFILE%\.dotnet\tools
            // Linux/Mac: ~/.dotnet/tools
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var globalToolsPath = Path.Combine(userProfile, ".dotnet", "tools");

            if (executablePath.StartsWith(globalToolsPath, StringComparison.OrdinalIgnoreCase))
                return InstallationType.Global;

            // Local tools are typically in the project's .config/dotnet-tools.json
            // and executed from the project directory
            return InstallationType.Local;
        }
        catch
        {
            return InstallationType.Unknown;
        }
    }

    /// <summary>
    /// Gets the update commands for the current installation type.
    /// Returns commands for both MCP Server and CLI packages.
    /// </summary>
    public static string GetUpdateCommand()
    {
        var installType = GetInstallationType();
        var flag = installType == InstallationType.Global ? " --global" : "";
        return $"dotnet tool update{flag} Sbroenne.ExcelMcp.McpServer && dotnet tool update{flag} Sbroenne.ExcelMcp.CLI";
    }

    /// <summary>
    /// Attempts to update both MCP Server and CLI tool packages.
    /// </summary>
    /// <returns>True if both updates succeeded, false otherwise.</returns>
    public static async Task<(bool Success, string Output)> TryUpdateAsync()
    {
        var installType = GetInstallationType();
        var flag = installType == InstallationType.Global ? " --global" : "";
        var packages = new[] { "Sbroenne.ExcelMcp.McpServer", "Sbroenne.ExcelMcp.CLI" };
        var allOutput = new List<string>();

        foreach (var package in packages)
        {
            var (success, output) = await RunUpdateCommandAsync($"tool update{flag} {package}");
            allOutput.Add($"{package}: {(success ? "OK" : "FAILED")} - {output}");
            if (!success)
                return (false, string.Join("\n", allOutput));
        }

        return (true, string.Join("\n", allOutput));
    }

    private static async Task<(bool Success, string Output)> RunUpdateCommandAsync(string arguments)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process == null)
                return (false, "Failed to start update process");

            var output = await process.StandardOutput.ReadToEndAsync();
            var error = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            return process.ExitCode == 0
                ? (true, output)
                : (false, error);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }
}

/// <summary>
/// Type of .NET tool installation.
/// </summary>
internal enum InstallationType
{
    Unknown,
    Global,
    Local
}


