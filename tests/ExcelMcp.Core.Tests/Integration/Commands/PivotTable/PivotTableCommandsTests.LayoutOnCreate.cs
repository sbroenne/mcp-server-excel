using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for PivotTable layout applied during create operations (Issue #366).
/// These tests verify that layout can be set immediately after PivotTable creation
/// in a single batch operation.
/// </summary>
public partial class PivotTableCommandsTests
{
    #region CreateFromRange with Layout

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateFromRange_WithTabularLayout_AppliesLayoutDuringCreate()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(CreateFromRange_WithTabularLayout_AppliesLayoutDuringCreate));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create pivot
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "LayoutPivot");
        Assert.True(createResult.Success, $"CreateFromRange failed: {createResult.ErrorMessage}");

        // Apply layout immediately after creation (same batch - simulates what MCP tool will do)
        var layoutResult = _pivotCommands.SetLayout(batch, "LayoutPivot", 1); // Tabular

        // Assert - Layout applied successfully
        Assert.True(layoutResult.Success, $"SetLayout failed: {layoutResult.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateFromRange_WithCompactLayout_AppliesLayoutDuringCreate()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(CreateFromRange_WithCompactLayout_AppliesLayoutDuringCreate));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create pivot and set layout
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "LayoutPivot");
        Assert.True(createResult.Success);

        var layoutResult = _pivotCommands.SetLayout(batch, "LayoutPivot", 0); // Compact

        // Assert
        Assert.True(layoutResult.Success, $"SetLayout failed: {layoutResult.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateFromRange_WithOutlineLayout_AppliesLayoutDuringCreate()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(CreateFromRange_WithOutlineLayout_AppliesLayoutDuringCreate));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create pivot and set layout
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "LayoutPivot");
        Assert.True(createResult.Success);

        var layoutResult = _pivotCommands.SetLayout(batch, "LayoutPivot", 2); // Outline

        // Assert
        Assert.True(layoutResult.Success, $"SetLayout failed: {layoutResult.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateFromRange_LayoutAndFields_BothApplyInSameBatch()
    {
        // Arrange - This test verifies the full workflow: create + layout + fields
        var testFile = CreateTestFileWithData(nameof(CreateFromRange_LayoutAndFields_BothApplyInSameBatch));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create pivot, set layout, add fields (all in same batch)
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "LayoutPivot");
        Assert.True(createResult.Success, $"CreateFromRange failed: {createResult.ErrorMessage}");

        var layoutResult = _pivotCommands.SetLayout(batch, "LayoutPivot", 1); // Tabular
        Assert.True(layoutResult.Success, $"SetLayout failed: {layoutResult.ErrorMessage}");

        var row1 = _pivotCommands.AddRowField(batch, "LayoutPivot", "Region");
        Assert.True(row1.Success, $"AddRowField Region failed: {row1.ErrorMessage}");

        var row2 = _pivotCommands.AddRowField(batch, "LayoutPivot", "Product");
        Assert.True(row2.Success, $"AddRowField Product failed: {row2.ErrorMessage}");

        var value = _pivotCommands.AddValueField(batch, "LayoutPivot", "Sales");
        Assert.True(value.Success, $"AddValueField failed: {value.ErrorMessage}");

        // Assert - Verify all operations succeeded
        var readResult = _pivotCommands.Read(batch, "LayoutPivot");
        Assert.True(readResult.Success);
        Assert.Equal("LayoutPivot", readResult.PivotTable!.Name);
    }

    #endregion

    #region CreateFromTable with Layout

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateFromTable_WithTabularLayout_AppliesLayoutDuringCreate()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(CreateFromTable_WithTabularLayout_AppliesLayoutDuringCreate));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create table first
        var tableCommands = new TableCommands();
        tableCommands.Create(batch, "SalesData", "SalesTable", "A1:D6", true, TableStylePresets.Medium2);

        // Act - Create pivot from table and set layout
        var createResult = _pivotCommands.CreateFromTable(
            batch, "SalesTable", "SalesData", "F1", "TableLayoutPivot");
        Assert.True(createResult.Success, $"CreateFromTable failed: {createResult.ErrorMessage}");

        var layoutResult = _pivotCommands.SetLayout(batch, "TableLayoutPivot", 1); // Tabular

        // Assert
        Assert.True(layoutResult.Success, $"SetLayout failed: {layoutResult.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateFromTable_WithOutlineLayout_AppliesLayoutDuringCreate()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(CreateFromTable_WithOutlineLayout_AppliesLayoutDuringCreate));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create table first
        var tableCommands = new TableCommands();
        tableCommands.Create(batch, "SalesData", "SalesTable", "A1:D6", true, TableStylePresets.Medium2);

        // Act - Create pivot from table and set layout
        var createResult = _pivotCommands.CreateFromTable(
            batch, "SalesTable", "SalesData", "F1", "TableLayoutPivot");
        Assert.True(createResult.Success);

        var layoutResult = _pivotCommands.SetLayout(batch, "TableLayoutPivot", 2); // Outline

        // Assert
        Assert.True(layoutResult.Success, $"SetLayout failed: {layoutResult.ErrorMessage}");
    }

    #endregion

    #region Persistence Tests

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateFromRange_LayoutAndFields_PersistsAfterSaveAndReopen()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(CreateFromRange_LayoutAndFields_PersistsAfterSaveAndReopen));

        // Act - Create, configure, and save
        using (var batch = ExcelSession.BeginBatch(testFile))
        {
            var createResult = _pivotCommands.CreateFromRange(
                batch, "SalesData", "A1:D6", "SalesData", "F1", "PersistPivot");
            Assert.True(createResult.Success);

            var layoutResult = _pivotCommands.SetLayout(batch, "PersistPivot", 1); // Tabular
            Assert.True(layoutResult.Success);

            var row = _pivotCommands.AddRowField(batch, "PersistPivot", "Region");
            Assert.True(row.Success);

            var value = _pivotCommands.AddValueField(batch, "PersistPivot", "Sales");
            Assert.True(value.Success);

            batch.Save();
        }

        // Assert - Reopen and verify
        using (var batch = ExcelSession.BeginBatch(testFile))
        {
            var listResult = _pivotCommands.List(batch);
            Assert.True(listResult.Success);
            Assert.Contains(listResult.PivotTables, pt => pt.Name == "PersistPivot");

            var fields = _pivotCommands.ListFields(batch, "PersistPivot");
            Assert.True(fields.Success);
            Assert.Contains(fields.Fields, f => f.Name == "Region");
        }
    }

    #endregion
}




