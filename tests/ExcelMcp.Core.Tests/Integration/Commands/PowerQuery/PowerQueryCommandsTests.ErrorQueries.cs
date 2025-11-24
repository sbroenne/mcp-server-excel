using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Regression tests for "Error Query" bug where List() suppresses exceptions
/// Issue: When COM error 0x800A03EC occurs, List() creates fake "Error Query {i}" entries
/// Expected: Exceptions should propagate naturally (CRITICAL-RULES.md Rule 1b)
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
[Trait("Speed", "Medium")]
public partial class PowerQueryCommandsTests
{
    /// <summary>
    /// Reproduces the "Error Query" bug with ConsumptionPlan_Base.xlsx
    /// LLM got result with 4 "Error Query" entries instead of actual query names
    /// </summary>
    [Fact]
    public void List_ConsumptionPlanBaseFile_ReturnsActualQueryNames()
    {
        // Arrange
        const string testFile = @"D:\source\mcp-server-excel\ConsumptionPlan_Base.xlsx";

        // Skip if file doesn't exist (not in all test environments)
        if (!System.IO.File.Exists(testFile))
        {
            return;
        }

        var dataModelCommands = new DataModelCommands();
        var commands = new PowerQueryCommands(dataModelCommands);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = commands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Queries);

        // Should have 4 actual queries, not "Error Query" entries
        Assert.Equal(4, result.Queries.Count);

        // Verify NO "Error Query" fake entries
        Assert.DoesNotContain(result.Queries, q => q.Name.StartsWith("Error Query", StringComparison.Ordinal));

        // Verify actual query names are present
        Assert.Contains(result.Queries, q => q.Name == "Milestones_Base");
        Assert.Contains(result.Queries, q => q.Name == "fnEnsureColumn");
        Assert.Contains(result.Queries, q => q.Name == "fnEnsureColumn_New");
        Assert.Contains(result.Queries, q => q.Name == "Milestones_Base_New");

        // Verify formulas are accessible (not empty due to catch block)
        foreach (var query in result.Queries)
        {
            Assert.NotEmpty(query.Formula);
            Assert.DoesNotContain("Error:", query.FormulaPreview);
        }
    }
}
