using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration;

/// <summary>
/// Guards the tool/operation counts the MCP server advertises in its own <c>--help</c> banner.
///
/// THE BUG THIS PREVENTS
/// ---------------------
/// The banner used to carry the hard-coded literal "Provides 22 tools with 195+ operations".
/// The real surface had grown to 31 tools / 326 operations, so the binary told users something
/// that contradicted every README, SKILL.md and the live <c>tools/list</c> response.
///
/// <see cref="McpToolSurface"/> now derives BOTH numbers by reflecting over the actual
/// <c>[McpServerToolType]</c>/<c>[McpServerTool]</c> registration, so the banner cannot drift
/// from the registration again. These tests lock in that derivation and cross-check it against
/// the ground truth enforced everywhere else (scripts/check-doc-counts.ps1).
/// </summary>
/// <inheritdoc/>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ToolSurface")]
[Trait("RequiresExcel", "false")]
public class McpToolSurfaceTests(ITestOutputHelper output)
{
    /// <summary>
    /// Ground truth, verified against the live <c>tools/list</c> response of the shipped server
    /// and enforced across every user-facing doc by scripts/check-doc-counts.ps1.
    /// </summary>
    private const int ExpectedToolCount = 31;

    /// <summary>Sum of the <c>action</c> enum values across all registered tools.</summary>
    private const int ExpectedOperationCount = 327;

    [Fact]
    public void ToolSurface_MatchesDocumentedGroundTruth()
    {
        foreach (var tool in McpToolSurface.Tools.OrderBy(t => t.Name, StringComparer.Ordinal))
        {
            output.WriteLine($"  {tool.Name}: {tool.OperationCount}");
        }

        Assert.Equal(ExpectedToolCount, McpToolSurface.ToolCount);
        Assert.Equal(ExpectedOperationCount, McpToolSurface.OperationCount);
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

    [Fact]
    public void HelpText_DoesNotCarryTheStaleHardCodedCounts()
    {
        var help = Program.BuildHelpText();

        Assert.DoesNotContain("22 tools", help, StringComparison.Ordinal);
        Assert.DoesNotContain("195+", help, StringComparison.Ordinal);
    }
}
