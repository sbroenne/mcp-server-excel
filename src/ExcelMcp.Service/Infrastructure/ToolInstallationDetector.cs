using System.Diagnostics;

namespace Sbroenne.ExcelMcp.Service.Infrastructure;

/// <summary>
/// Detects whether the CLI is installed as a global or local .NET tool.
/// </summary>
public static class ToolInstallationDetector
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
    /// Gets the update command for the current installation type.
    /// </summary>
    public static string GetUpdateCommand()
    {
        var installType = GetInstallationType();
        return installType switch
        {
            InstallationType.Global => "dotnet tool update --global Sbroenne.ExcelMcp.McpServer",
            InstallationType.Local => "dotnet tool update Sbroenne.ExcelMcp.McpServer",
            _ => "dotnet tool update --global Sbroenne.ExcelMcp.McpServer"
        };
    }

    /// <summary>
    /// Attempts to update the CLI tool.
    /// </summary>
    /// <returns>True if update succeeded, false otherwise.</returns>
    public static async Task<(bool Success, string Output)> TryUpdateAsync()
    {
        try
        {
            var command = GetUpdateCommand();
            var parts = command.Split(' ', 2);
            var fileName = parts[0];
            var arguments = parts.Length > 1 ? parts[1] : string.Empty;

            var psi = new ProcessStartInfo
            {
                FileName = fileName,
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

            if (process.ExitCode == 0)
                return (true, output);

            return (false, error);
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
public enum InstallationType
{
    Unknown,
    Global,
    Local
}


