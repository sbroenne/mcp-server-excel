using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests verifying List() handles manually created tables (ListObjects without QueryTables)
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
[Trait("Speed", "Medium")]
public partial class PowerQueryCommandsTests
{
    /// <summary>
    /// Verifies List() handles workbooks with both Power Queries AND manually created tables
    /// Manually created tables don't have QueryTable property and throw 0x800A03EC
    /// List() should skip those tables gracefully without creating "Error Query" entries
    /// </summary>
    [Fact]
    public void List_WorkbookWithManualTable_ReturnsOnlyQueries()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        var dataModelCommands = new DataModelCommands();
        var commands = new PowerQueryCommands(dataModelCommands);

        const string queryName = "TestQuery";
        const string mCode = @"let
    Source = #table(
        {""Column1"", ""Column2""},
        {
            {""A"", ""B""},
            {""C"", ""D""}
        }
    )
in
    Source";

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);

        // Step 1: Create a manually created table (no Power Query connection)
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? listObjects = null;
            try
            {
                sheet = ctx.Book.Worksheets.Item(1);
                sheet.Name = "TestSheet";

                // Add some data
                range = sheet.Range["A1:B3"];
                range.Value2 = new object[,]
                {
                    { "Header1", "Header2" },
                    { "Data1", "Data2" },
                    { "Data3", "Data4" }
                };

                // Create a manual table (ListObject) - NO QueryTable
                listObjects = sheet.ListObjects;
                dynamic? listObject = listObjects.Add(
                    1,              // xlSrcRange (manual table from range)
                    range,          // Source range
                    Type.Missing,   // LinkSource
                    1,              // xlYes (has headers)
                    Type.Missing    // Destination
                );
                listObject.Name = "ManualTable";
                ComUtilities.Release(ref listObject!);
            }
            finally
            {
                ComUtilities.Release(ref listObjects!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });

        // Step 2: Create a Power Query (connection-only)
        commands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Step 3: List queries
        var result = commands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Queries);

        // Should have 1 query (not "Error Query" entries from manual table)
        Assert.Single(result.Queries);

        // Verify NO "Error Query" fake entries
        Assert.DoesNotContain(result.Queries, q => q.Name.StartsWith("Error Query", StringComparison.Ordinal));

        // Verify the actual query is present
        var query = Assert.Single(result.Queries);
        Assert.Equal(queryName, query.Name);
        Assert.NotEmpty(query.Formula);
        Assert.DoesNotContain("Error:", query.FormulaPreview);

        // Query should be connection-only (manual table shouldn't affect this)
        Assert.True(query.IsConnectionOnly);
    }
}
