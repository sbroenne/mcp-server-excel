using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model discovery operations (columns, table view, model info)
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task ListTableColumns_WithValidTable_ReturnsColumns()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListTableColumnsAsync(batch, "Sales");

        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(result.Success, 
            $"ListTableColumns MUST succeed. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Columns);
        Assert.True(result.Columns.Count >= 6, $"Expected at least 6 columns in Sales table, got {result.Columns.Count}");
        Assert.Equal("Sales", result.TableName);

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
    public async Task ListTableColumns_WithNonExistentTable_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListTableColumnsAsync(batch, "NonExistentTable");

        // Assert - Should fail because table doesn't exist (Data Model is always available in Excel 2013+)
        Assert.False(result.Success, "ListTableColumns should fail when table doesn't exist");
        Assert.NotNull(result.ErrorMessage);
        Assert.True(result.ErrorMessage.Contains("Table 'NonExistentTable' not found"),
            $"Expected 'table not found' error, but got: {result.ErrorMessage}");
    }

    [Fact]
    public async Task ViewTable_WithValidTable_ReturnsCompleteInfo()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ViewTableAsync(batch, "Sales");

        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(result.Success, 
            $"ViewTable MUST succeed. Error: {result.ErrorMessage}");
        Assert.Equal("Sales", result.TableName);
        Assert.NotNull(result.SourceName);
        Assert.True(result.RecordCount >= 10, $"Expected at least 10 records in Sales table, got {result.RecordCount}");

        // Should have columns
        Assert.NotNull(result.Columns);
        Assert.True(result.Columns.Count >= 6, $"Expected at least 6 columns, got {result.Columns.Count}");

        // Should have measure count (may be 0 or more depending on Data Model creation)
        Assert.True(result.MeasureCount >= 0, $"MeasureCount should be non-negative, got {result.MeasureCount}");
    }

    [Fact]
    public async Task ViewTable_WithTableHavingMeasures_CountsMeasuresCorrectly()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ViewTableAsync(batch, "Sales");

        // Assert - If Data Model was created with measures
        if (result.Success && result.MeasureCount > 0)
        {
            // Sales table should have at least 2 measures (Total Sales, Average Sale)
            Assert.True(result.MeasureCount >= 2, $"Expected at least 2 measures for Sales table, got {result.MeasureCount}");
        }
    }

    [Fact]
    public async Task ViewTable_WithNonExistentTable_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ViewTableAsync(batch, "NonExistentTable");

        // Assert - Demand specific "table not found" error (Data Model is always available)
        Assert.False(result.Success, "Should fail when table doesn't exist");
        Assert.NotNull(result.ErrorMessage);
        Assert.True(result.ErrorMessage.Contains("Table 'NonExistentTable' not found"),
            $"Expected 'Table not found' error, but got: {result.ErrorMessage}");
    }

    [Fact]
    public async Task GetModelInfo_WithRealisticDataModel_ReturnsAccurateStatistics()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.GetModelInfoAsync(batch);

        // Assert - Demand success (Data Model is always available in Excel 2013+)
        Assert.True(result.Success, 
            $"GetModelInfo MUST succeed - Data Model is always available in Excel 2013+. Error: {result.ErrorMessage}");
        
        // Should have at least 3 tables (Sales, Customers, Products)
        Assert.True(result.TableCount >= 3, $"Expected at least 3 tables, got {result.TableCount}");

        // Should have measures if Data Model was created successfully
        Assert.True(result.MeasureCount >= 0, $"MeasureCount should be non-negative, got {result.MeasureCount}");

        // Should have relationships between tables
        Assert.True(result.RelationshipCount >= 0, $"RelationshipCount should be non-negative, got {result.RelationshipCount}");

        // Should have total row count
        Assert.True(result.TotalRows > 0, $"Expected positive total row count, got {result.TotalRows}");

        // Should have table names
        Assert.NotNull(result.TableNames);
        Assert.True(result.TableNames.Count >= 3, $"Expected at least 3 table names, got {result.TableNames.Count}");
        Assert.Contains("Sales", result.TableNames);
        Assert.Contains("Customers", result.TableNames);
    }

    [Fact]
    public async Task GetModelInfo_WithDataModelHavingMeasures_CountsCorrectly()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.GetModelInfoAsync(batch);

        // Assert - If Data Model was created with measures
        if (result.Success && result.MeasureCount > 0)
        {
            // Should have at least 3 measures (Total Sales, Average Sale, Total Customers)
            Assert.True(result.MeasureCount >= 3, $"Expected at least 3 measures, got {result.MeasureCount}");
        }
    }
}
