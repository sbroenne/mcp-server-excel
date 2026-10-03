using System.Text.Json;
using System.Text.RegularExpressions;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Verifies authored calculation examples use the server's published names and actions.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
[Trait("RequiresExcel", "false")]
public sealed class CalculationGuidanceContractTests : McpIntegrationTestBase
{
    public CalculationGuidanceContractTests(ITestOutputHelper output)
        : base(output, "CalculationGuidanceContractClient")
    {
    }

    [Fact]
    public async Task AuthoredCalculationExamples_MatchGeneratedToolSchema()
    {
        var tools = await Client!.ListToolsAsync(cancellationToken: TestCancellationToken);
        var calculationTool = tools.Single(tool => tool.Name == "calculation_mode");
        var properties = calculationTool.JsonSchema.GetProperty("properties");
        var schemaParameterNames = properties.EnumerateObject()
            .Select(property => property.Name)
            .ToHashSet(StringComparer.Ordinal);
        var actionNames = properties.GetProperty("action").GetProperty("enum")
            .EnumerateArray()
            .Select(value => value.GetString())
            .Where(value => value != null)
            .Select(value => value!)
            .ToHashSet(StringComparer.Ordinal);
        var requiredParameterNames = GetRequiredPropertyNames(calculationTool.JsonSchema);
        var repoRoot = FindRepoRoot();
        var sourcePaths = Directory.GetFiles(Path.Combine(repoRoot, "docs", "reference"), "*.md");
        var exampleCount = 0;

        foreach (var sourcePath in sourcePaths)
        {
            var content = await File.ReadAllTextAsync(
                sourcePath,
                TestCancellationToken);
            var examples = Regex.Matches(content, @"\bcalculation_mode\((?<arguments>[^)\r\n]+)\)");

            foreach (Match example in examples)
            {
                exampleCount++;
                var arguments = example.Groups["arguments"].Value.Split(
                    ',',
                    StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                var parameterNames = new HashSet<string>(StringComparer.Ordinal);
                string? action = null;

                foreach (var argument in arguments)
                {
                    var namedArgument = Regex.Match(
                        argument,
                        @"^(?<name>[a-z][a-z0-9_]*)\s*:\s*(?:['""](?<value>[^'""]+)['""]|(?<boolean>true|false)|(?<variable>[a-zA-Z_][a-zA-Z0-9_]*))$");
                    Assert.True(
                        namedArgument.Success,
                        $"Use named MCP arguments in `{example.Value}` from {sourcePath}.");

                    var parameterName = namedArgument.Groups["name"].Value;
                    Assert.Contains(parameterName, schemaParameterNames);
                    if (namedArgument.Groups["boolean"].Success)
                    {
                        var type = properties.GetProperty(parameterName).GetProperty("type");
                        Assert.True(type.ValueKind == JsonValueKind.Array
                            ? type.EnumerateArray().Any(value => value.GetString() == "boolean")
                            : type.GetString() == "boolean");
                    }
                    if (namedArgument.Groups["variable"].Success)
                    {
                        Assert.True(parameterName is "session_id" or "mode",
                            $"Only session_id and a remembered mode may use workflow variables: {argument}");
                    }
                    parameterNames.Add(parameterName);
                    if (parameterName == "action")
                    {
                        action = namedArgument.Groups["value"].Value;
                    }
                }

                Assert.NotNull(action);
                Assert.Contains(action, actionNames);
                foreach (var requiredParameterName in requiredParameterNames)
                {
                    Assert.Contains(requiredParameterName, parameterNames);
                }
            }
        }
        Assert.True(exampleCount > 0, "No calculation examples were checked.");
    }

    private static HashSet<string> GetRequiredPropertyNames(JsonElement schema)
    {
        if (!schema.TryGetProperty("required", out var required))
        {
            return [];
        }

        return new HashSet<string>(
            required.EnumerateArray()
                .Select(value => value.GetString())
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .Select(static value => value!),
            StringComparer.Ordinal);
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
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
