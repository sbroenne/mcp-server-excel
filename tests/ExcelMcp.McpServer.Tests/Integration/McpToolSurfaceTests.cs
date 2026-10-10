using System.Reflection;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration;

/// <summary>
/// Guards the tool/operation counts the MCP server advertises in its own <c>--help</c> banner.
///
/// THE BUG THIS PREVENTS
/// ---------------------
/// The banner used to carry the hard-coded literal "Provides 22 tools with 195+ operations".
/// The real surface had grown to 31 command categories / 326 operations, so the binary told users
/// something that contradicted every README, SKILL.md and the live <c>tools/list</c> response.
///
/// <see cref="McpToolSurface"/> derives the banner numbers from the actual
/// <c>[McpServerToolType]</c>/<c>[McpServerTool]</c> registration. Read-only action groups add
/// MCP endpoints without changing the CLI command-category count.
/// </summary>
/// <inheritdoc/>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ToolSurface")]
[Trait("Feature", "GeneratedContracts")]
[Trait("RequiresExcel", "false")]
public class McpToolSurfaceTests(ITestOutputHelper output)
{
    [Fact]
    public void ToolSurface_MatchesDocumentedGroundTruth()
    {
        foreach (var tool in McpToolSurface.Tools.OrderBy(t => t.Name, StringComparer.Ordinal))
        {
            output.WriteLine($"  {tool.Name}: {tool.OperationCount}");
        }

        var expected = _CliCategoryMetadata.ValidActionsByCommand
            .Where(pair => pair.Key != "diag")
            .ToDictionary(pair => pair.Key, pair => pair.Value.Count, StringComparer.Ordinal);
        expected.Add("file", Enum.GetValues<FileAction>().Length);
        var expectedReadTools = typeof(IRangeCommands).Assembly.GetTypes()
            .Where(type => type.GetCustomAttribute<McpReadOnlyActionsAttribute>() is not null &&
                type.GetCustomAttribute<McpToolAttribute>() is not null)
            .Select(type => $"{type.GetCustomAttribute<McpToolAttribute>()!.ToolName}_read")
            .Concat(new[] { typeof(ExcelFileTool), typeof(ExcelWorksheetTool) }
                .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Static))
                .Select(method => method.GetCustomAttribute<McpServerToolAttribute>())
                .Where(attribute => attribute?.Name?.EndsWith("_read", StringComparison.Ordinal) == true)
                .Select(attribute => attribute!.Name!))
            .ToHashSet(StringComparer.Ordinal);
        Assert.Equal(expected.Count + expectedReadTools.Count, McpToolSurface.ToolCount);
        Assert.True(expectedReadTools.SetEquals(McpToolSurface.Tools
            .Where(tool => tool.Name.EndsWith("_read", StringComparison.Ordinal))
            .Select(tool => tool.Name)));
        Assert.Equal(expected.Values.Sum(), McpToolSurface.OperationCount);
        var groupedTools = McpToolSurface.Tools
            .GroupBy(tool => tool.Name.EndsWith("_read", StringComparison.Ordinal)
                ? tool.Name[..^"_read".Length]
                : tool.Name)
            .ToDictionary(group => group.Key, group => group.Sum(tool => tool.OperationCount), StringComparer.Ordinal);
        foreach (var (name, operationCount) in groupedTools)
        {
            var command = name switch
            {
                "worksheet" => "sheet",
                _ => name.Replace("_", "", StringComparison.Ordinal)
            };
            Assert.Equal(expected[command], operationCount);
        }
    }

    [Fact]
    public void EveryRegisteredTool_ExposesAnActionEnum()
    {
        // The operation count is only trustworthy while every tool routes through an `action`
        // enum. A tool without one would silently contribute 0 operations.
        var offenders = McpToolSurface.Tools.Where(t => t.OperationCount == 0).ToList();

        Assert.True(
            offenders.Count == 0,
            "These MCP tools have no 'action' enum parameter, so McpToolSurface cannot count their " +
            $"operations: {string.Join(", ", offenders.Select(t => t.Name))}");
    }

    [Fact]
    public void ToolNames_AreUnique()
    {
        var duplicates = McpToolSurface.Tools
            .GroupBy(t => t.Name, StringComparer.Ordinal)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToList();

        Assert.True(duplicates.Count == 0, $"Duplicate MCP tool names: {string.Join(", ", duplicates)}");
    }

    [Fact]
    public void HelpText_AdvertisesTheDerivedCounts()
    {
        var help = Program.BuildHelpText();
        output.WriteLine(help);

        Assert.Contains(
            $"Provides {McpToolSurface.ToolCount} tools with {McpToolSurface.OperationCount} operations",
            help,
            StringComparison.Ordinal);
    }

}
