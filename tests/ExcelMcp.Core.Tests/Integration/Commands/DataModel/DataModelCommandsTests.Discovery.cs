using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model discovery operations (columns, table view, model info)
/// Uses shared Data Model file from fixture.
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task ListTableColumns_WithValidTable_ReturnsColumns()
    {
        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.ListTableColumnsAsync(batch, "SalesTable");

        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(result.Success,
            $"ListTableColumns MUST succeed. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.True(result.Columns.Count >= 6, $"Expected at least 6 columns in SalesTable, got {result.Columns.Count}");
        Assert.Equal("SalesTable", result.TableName);

        // Verify expected columns exist
        var columnNames = result.Columns.Select(c => c.Name).ToList();
        Assert.Contains("SalesID", columnNames);
        Assert.Contains("CustomerID", columnNames);
        Assert.Contains("Amount", columnNames);

        // Verify column properties
        var amountColumn = result.Columns.FirstOrDefault(c => c.Name == "Amount");
        if (amountColumn != null)
        {
            Assert.NotNull(amountColumn.DataType);
            Assert.False(amountColumn.IsCalculated); // Should be a data column, not calculated
        }
    }

    [Fact]
    public async Task ViewTable_WithValidTable_ReturnsCompleteInfo()
    {
        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.ViewTableAsync(batch, "SalesTable");

        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(result.Success,
            $"ViewTable MUST succeed. Error: {result.ErrorMessage}");
        Assert.Equal("SalesTable", result.TableName);
        Assert.NotNull(result.SourceName);
        Assert.True(result.RecordCount >= 10, $"Expected at least 10 records in Sales table, got {result.RecordCount}");

        // Should have columns
        Assert.NotNull(result.Columns);
        Assert.True(result.Columns.Count >= 6, $"Expected at least 6 columns, got {result.Columns.Count}");

        // Should have measure count (from fixture creation)
        Assert.True(result.MeasureCount >= 0, $"MeasureCount should be non-negative, got {result.MeasureCount}");
    }

    [Fact]
    public async Task ViewTable_WithTableHavingMeasures_CountsMeasuresCorrectly()
    {
        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.ViewTableAsync(batch, "SalesTable");

        // Assert - Fixture created 3 measures on SalesTable
        Assert.True(result.Success, $"ViewTable failed: {result.ErrorMessage}");
        Assert.True(result.MeasureCount >= 2, $"Expected at least 2 measures for SalesTable, got {result.MeasureCount}");
    }

    [Fact]
    public async Task GetModelInfo_WithRealisticDataModel_ReturnsAccurateStatistics()
    {
        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.GetModelInfoAsync(batch);

        // Assert - Demand success (Data Model is always available in Excel 2013+)
        Assert.True(result.Success,
            $"GetModelInfo MUST succeed - Data Model is always available in Excel 2013+. Error: {result.ErrorMessage}");

        // Should have exactly 3 tables (from fixture)
        Assert.Equal(3, result.TableCount);

        // Should have exactly 3 measures (from fixture)
        Assert.Equal(3, result.MeasureCount);

        // Should have exactly 2 relationships (from fixture)
        Assert.Equal(2, result.RelationshipCount);

        // Should have total row count
        Assert.True(result.TotalRows > 0, $"Expected positive total row count, got {result.TotalRows}");

        // Should have table names
        Assert.NotNull(result.TableNames);
        Assert.Equal(3, result.TableNames.Count);
        Assert.Contains("SalesTable", result.TableNames);
        Assert.Contains("CustomersTable", result.TableNames);
        Assert.Contains("ProductsTable", result.TableNames);
    }
}
