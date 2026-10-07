using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration;

/// <summary>
/// Keeps <c>.github/usage-analytics-weights.json</c> aligned with the registered MCP tools.
/// The public usage report multiplies every action by its hand-picked work level; an action
/// missing from the file would silently drop out of the "share of work" figures.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "UsageAnalytics")]
[Trait("RequiresExcel", "false")]
public class UsageAnalyticsWeightsTests
{
    private const string ReadSuffix = "_read";

    [Fact]
    public void WeightsFile_ListsEveryMcpToolActionExactlyOnce()
    {
        using var weights = LoadWeights();
        var tools = weights.RootElement.GetProperty("tools");
        var expected = DiscoverMcpActions();

        var listedMcpTools = tools.EnumerateObject()
            .Where(tool => !IsCliOnly(tool.Value))
            .Select(tool => tool.Name)
            .ToHashSet(StringComparer.Ordinal);
        var missingTools = expected.Keys.Except(listedMcpTools).Order().ToList();
        var staleTools = listedMcpTools.Except(expected.Keys).Order().ToList();
        Assert.True(missingTools.Count == 0,
            $"Add these MCP tools to .github/usage-analytics-weights.json: {string.Join(", ", missingTools)}");
        Assert.True(staleTools.Count == 0,
            $"Remove these tools from .github/usage-analytics-weights.json or mark them cliOnly: {string.Join(", ", staleTools)}");

        var problems = new List<string>();
        foreach (var (toolName, actions) in expected)
        {
            var listedActions = tools.GetProperty(toolName).GetProperty("actions").EnumerateObject()
                .Select(action => action.Name)
                .ToHashSet(StringComparer.Ordinal);
            problems.AddRange(actions.Except(listedActions).Order().Select(action => $"missing {toolName}/{action}"));
            problems.AddRange(listedActions.Except(actions).Order().Select(action => $"stale {toolName}/{action}"));
        }

        Assert.True(problems.Count == 0,
            "Update .github/usage-analytics-weights.json: " + string.Join("; ", problems));
    }

    [Fact]
    public void WeightsFile_UsesOnlyDeclaredLevelsAndFeatures()
    {
        using var weights = LoadWeights();
        var root = weights.RootElement;
        var levels = root.GetProperty("levels").EnumerateObject().ToDictionary(
            level => level.Name, level => level.Value.GetInt32(), StringComparer.Ordinal);
        var features = root.GetProperty("features").EnumerateArray()
            .Select(feature => feature.GetString()!)
            .ToHashSet(StringComparer.Ordinal);

        Assert.NotEmpty(levels);
        Assert.All(levels, level => Assert.True(level.Value > 0, $"Level '{level.Key}' must be positive."));
        Assert.Equal(levels.Count, levels.Values.Distinct().Count());
        Assert.Contains("other", features);

        var problems = new List<string>();
        foreach (var tool in root.GetProperty("tools").EnumerateObject())
        {
            var feature = tool.Value.GetProperty("feature").GetString();
            if (feature is null || !features.Contains(feature))
            {
                problems.Add($"{tool.Name} uses undeclared feature '{feature}'");
            }

            var actions = tool.Value.GetProperty("actions").EnumerateObject().ToList();
            if (actions.Count == 0)
            {
                problems.Add($"{tool.Name} lists no actions");
            }

            problems.AddRange(actions
                .Where(action => action.Value.GetString() is not { } level || !levels.ContainsKey(level))
                .Select(action => $"{tool.Name}/{action.Name} uses undeclared level '{action.Value}'"));
        }

        foreach (var excluded in root.GetProperty("excludedActions").EnumerateArray().Select(item => item.GetString()!))
        {
            var parts = excluded.Split('/');
            if (parts.Length != 2 ||
                !root.GetProperty("tools").TryGetProperty(parts[0], out var tool) ||
                !tool.GetProperty("actions").TryGetProperty(parts[1], out _))
            {
                problems.Add($"excluded action '{excluded}' is not a listed tool action");
            }
        }

        Assert.True(problems.Count == 0, string.Join("; ", problems));
    }

    /// <summary>
    /// Returns MCP tool actions grouped by the tool name telemetry reports after removing the
    /// <c>_read</c> suffix of read-only endpoints.
    /// </summary>
    private static Dictionary<string, HashSet<string>> DiscoverMcpActions()
    {
        const BindingFlags MethodFlags = BindingFlags.Public | BindingFlags.NonPublic |
                                         BindingFlags.Static | BindingFlags.Instance |
                                         BindingFlags.DeclaredOnly;
        var result = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        var toolAssembly = typeof(McpToolSurface).Assembly;
        foreach (var toolType in toolAssembly.GetTypes()
                     .Where(type => type.GetCustomAttribute<McpServerToolTypeAttribute>() is not null))
        {
            foreach (var method in toolType.GetMethods(MethodFlags))
            {
                if (method.GetCustomAttribute<McpServerToolAttribute>() is not { } toolAttribute)
                {
                    continue;
                }

                var toolName = toolAttribute.Name ?? method.Name;
                if (toolName.EndsWith(ReadSuffix, StringComparison.Ordinal))
                {
                    toolName = toolName[..^ReadSuffix.Length];
                }

                var actionParameter = method.GetParameters().Single(parameter => parameter.Name == "action");
                var actionType = Nullable.GetUnderlyingType(actionParameter.ParameterType) ?? actionParameter.ParameterType;
                Assert.True(actionType.IsEnum, $"{toolName} must expose an action enum.");
                if (!result.TryGetValue(toolName, out var actions))
                {
                    actions = new HashSet<string>(StringComparer.Ordinal);
                    result.Add(toolName, actions);
                }

                foreach (var field in actionType.GetFields(BindingFlags.Public | BindingFlags.Static))
                {
                    var name = field.GetCustomAttribute<JsonStringEnumMemberNameAttribute>()?.Name;
                    Assert.False(string.IsNullOrWhiteSpace(name),
                        $"{actionType.Name}.{field.Name} has no wire action name.");
                    actions.Add(name!);
                }
            }
        }

        return result;
    }

    private static bool IsCliOnly(JsonElement tool) =>
        tool.TryGetProperty("cliOnly", out var cliOnly) && cliOnly.GetBoolean();

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
