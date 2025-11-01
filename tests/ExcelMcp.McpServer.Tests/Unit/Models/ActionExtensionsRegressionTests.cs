using Xunit;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit.Models;

/// <summary>
/// CRITICAL REGRESSION TEST: Ensure all enum values are mapped in ToActionString()
/// Bug: Missing enum mappings caused ArgumentException instead of JSON error response
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("BugFix", "EnumMapping")]
public class ActionExtensionsRegressionTests
{
    [Theory]
    [InlineData(RangeAction.GetUsedRange, "get-used-range")]
    [InlineData(RangeAction.GetCurrentRegion, "get-current-region")]
    [InlineData(RangeAction.GetRangeInfo, "get-range-info")]
    [InlineData(RangeAction.ClearFormats, "clear-formats")]
    [InlineData(RangeAction.CopyFormulas, "copy-formulas")]
    [InlineData(RangeAction.InsertCells, "insert-cells")]
    [InlineData(RangeAction.DeleteCells, "delete-cells")]
    [InlineData(RangeAction.InsertRows, "insert-rows")]
    [InlineData(RangeAction.DeleteRows, "delete-rows")]
    [InlineData(RangeAction.InsertColumns, "insert-columns")]
    [InlineData(RangeAction.DeleteColumns, "delete-columns")]
    [InlineData(RangeAction.ListHyperlinks, "list-hyperlinks")]
    [InlineData(RangeAction.GetHyperlink, "get-hyperlink")]
    [InlineData(RangeAction.FormatRange, "format-range")]
    [InlineData(RangeAction.ValidateRange, "validate-range")]
    public void RangeAction_ToActionString_PreviouslyMissingMappings(RangeAction action, string expected)
    {
        // Act
        var result = action.ToActionString();

        // Assert
        Assert.Equal(expected, result);
    }
}
