using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Version checking tool for MCP server.
/// Allows checking if a newer version of ExcelMcp is available on NuGet.org.
///
/// LLM Usage Pattern:
/// - Use "check" action to verify if updates are available
/// - Inform users about available updates and how to upgrade
/// </summary>
[McpServerToolType]
public static class ExcelVersionTool
{
    private const string PackageId = "Sbroenne.ExcelMcp.McpServer";

    /// <summary>
    /// Check for available updates on NuGet.org
    /// </summary>
    [McpServerTool(Name = "excel_version")]
    [Description("Check for ExcelMcp updates on NuGet.org. Supports: check.")]
    public static async Task<string> ExcelVersion(
        [Description("Action to perform: check")]
        string action)
    {
        try
        {
            switch (action.ToLowerInvariant())
            {
                case "check":
                    return await CheckVersion();

                default:
                    throw new ModelContextProtocol.McpException($"Unknown action '{action}'. Supported: check");
            }
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            throw new ModelContextProtocol.McpException($"Version check failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Checks if a newer version is available on NuGet.org
    /// </summary>
    private static async Task<string> CheckVersion()
    {
        var checker = new VersionChecker();
        var result = await checker.CheckForUpdatesAsync(PackageId);

        if (!result.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                error = result.ErrorMessage,
                currentVersion = result.CurrentVersion,
                message = "Failed to check for updates",
                suggestedNextActions = new[]
                {
                    "Check your internet connection",
                    "Try again later",
                    "Visit https://www.nuget.org/packages/Sbroenne.ExcelMcp.McpServer for manual version check"
                },
                workflowHint = "Version check failed. You may still use the current version."
            }, ExcelToolsBase.JsonOptions);
        }

        if (result.IsOutdated)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                isOutdated = true,
                currentVersion = result.CurrentVersion,
                latestVersion = result.LatestVersion,
                packageId = result.PackageId,
                message = $"A newer version ({result.LatestVersion}) is available. You are running version {result.CurrentVersion}.",
                updateCommand = $"dotnet tool update -g {result.PackageId}",
                suggestedNextActions = new[]
                {
                    $"Update with: dotnet tool update -g {result.PackageId}",
                    "Restart VS Code after updating (this will restart the MCP server)",
                    "Verify update with excel_version check action"
                },
                workflowHint = "Update available. Run the update command and restart VS Code to use the latest version."
            }, ExcelToolsBase.JsonOptions);
        }

        return JsonSerializer.Serialize(new
        {
            success = true,
            isOutdated = false,
            currentVersion = result.CurrentVersion,
            latestVersion = result.LatestVersion,
            packageId = result.PackageId,
            message = $"You are running the latest version ({result.CurrentVersion}).",
            suggestedNextActions = new[]
            {
                "Continue using ExcelMcp tools",
                "Check for updates periodically",
                "Visit https://github.com/sbroenne/mcp-server-excel for documentation"
            },
            workflowHint = "You're up to date. Continue your Excel automation workflow."
        }, ExcelToolsBase.JsonOptions);
    }
}
