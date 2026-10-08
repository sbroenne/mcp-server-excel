using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

/// <summary>
/// Keeps <c>.github/usage-analytics-weights.json</c> aligned with every operation name excelcli
/// reports. The public usage report maps CLI categories to MCP tools through
/// <c>cliCategories</c>; an unmapped name would silently drop out of the "share of work" figures.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class UsageAnalyticsWeightsTests
{
    [Fact]
    public void WeightsFile_MapsEveryReportedCliOperationToExactlyOneWeightedAction()
    {
        using var weights = LoadWeights();
        var root = weights.RootElement;
        var tools = root.GetProperty("tools");
        var cliCategories = root.GetProperty("cliCategories");
        var reported = ReportedCliOperations();
        var problems = new List<string>();
        var usedToolActions = new HashSet<string>(StringComparer.Ordinal);

        foreach (var (category, action) in reported)
        {
            if (!cliCategories.TryGetProperty(category, out var targets))
            {
                problems.Add($"cliCategories has no entry for '{category}'");
                continue;
            }

            var matches = targets.EnumerateArray()
                .Select(target => target.GetString()!)
                .Where(target => tools.TryGetProperty(target, out var tool) &&
                                 tool.GetProperty("actions").TryGetProperty(action, out _))
                .ToList();
            if (matches.Count != 1)
            {
                problems.Add($"{category}/{action} matches {matches.Count} weighted tool actions");
                continue;
            }

            usedToolActions.Add($"{matches[0]}/{action}");
        }

        var reportedCategories = reported.Select(operation => operation.Category).ToHashSet(StringComparer.Ordinal);
        problems.AddRange(cliCategories.EnumerateObject()
            .Where(category => !reportedCategories.Contains(category.Name))
            .Select(category => $"stale cliCategories entry '{category.Name}'"));

        foreach (var tool in tools.EnumerateObject()
                     .Where(tool => tool.Value.TryGetProperty("cliOnly", out var cliOnly) && cliOnly.GetBoolean()))
        {
            problems.AddRange(tool.Value.GetProperty("actions").EnumerateObject()
                .Where(action => !usedToolActions.Contains($"{tool.Name}/{action.Name}"))
                .Select(action => $"stale CLI-only action {tool.Name}/{action.Name}"));
        }

        Assert.True(problems.Count == 0,
            "Update .github/usage-analytics-weights.json: " + string.Join("; ", problems));
    }

    /// <summary>
    /// Every category/action pair <see cref="CliTelemetry.ResolveOperation"/> can report.
    /// </summary>
    private static List<(string Category, string Action)> ReportedCliOperations()
    {
        var builtInField = typeof(CliTelemetry).GetField("BuiltInActions", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(builtInField);
        var builtIns = Assert.IsType<Dictionary<string, string[]>>(builtInField.GetValue(null));

        var commands = builtIns
            .SelectMany(pair => pair.Value.Select(action => $"{pair.Key}.{action}"))
            .Concat(ServiceRegistry.ValidActionsByCategory
                .SelectMany(pair => pair.Value.Select(action => $"{pair.Key}.{action}")))
            .Append("unknown-command.unknown-action");

        var operations = commands
            .Select(CliTelemetry.ResolveOperation)
            .Distinct()
            .ToList();
        Assert.Contains((CliTelemetry.UnknownOperationPart, CliTelemetry.UnknownOperationPart), operations);
        Assert.Contains(("range", "get-values"), operations);
        return operations;
    }

    private static JsonDocument LoadWeights() =>
        JsonDocument.Parse(File.ReadAllText(Path.Combine(FindRepoRoot(), ".github", "usage-analytics-weights.json")));

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not find repository root.");
    }
}
