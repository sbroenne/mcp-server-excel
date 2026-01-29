using System.Reflection;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Displays the current excelcli version. Use --check to check for updates.
/// </summary>
internal sealed class VersionCommand : AsyncCommand<VersionCommand.Settings>
{
    private readonly ICliConsole _console;

    public VersionCommand(ICliConsole console)
    {
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        var currentVersion = GetCurrentVersion();

        if (settings.Check)
        {
            var latestVersion = await NuGetVersionChecker.GetLatestVersionAsync(cancellationToken);
            var updateAvailable = latestVersion != null && CompareVersions(currentVersion, latestVersion) < 0;

            _console.WriteJson(new
            {
                currentVersion,
                latestVersion = latestVersion ?? "unknown",
                updateAvailable,
                updateCommand = updateAvailable ? "dotnet tool update --global Sbroenne.ExcelMcp.CLI" : null,
                releaseNotesUrl = updateAvailable ? "https://github.com/sbroenne/mcp-server-excel/releases/latest" : null
            });
        }
        else
        {
            VersionReporter.WriteVersion();
        }

        return 0;
    }

    private static string GetCurrentVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        // Strip git hash suffix (e.g., "1.2.0+abc123" -> "1.2.0")
        return informational?.Split('+')[0] ?? assembly.GetName().Version?.ToString() ?? "0.0.0";
    }

    private static int CompareVersions(string current, string latest)
    {
        if (Version.TryParse(current, out var currentVer) && Version.TryParse(latest, out var latestVer))
            return currentVer.CompareTo(latestVer);
        return string.Compare(current, latest, StringComparison.Ordinal);
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandOption("--check")]
        public bool Check { get; init; }
    }
}
