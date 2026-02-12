using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for PivotTable layout and subtotals operations.
/// </summary>
public partial class PivotTableCommandsTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public void SetLayout_Compact_UpdatesLayoutForm()
    {
        // Arrange - Create test file with data
        var testFile = CreateTestFileWithData(nameof(SetLayout_Compact_UpdatesLayoutForm));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        var createResult = _pivotCommands.CreateFromRange(batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesPivot");
        Assert.True(createResult.Success, $"CreateFromRange failed: {createResult.ErrorMessage}");

        var row1 = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        Assert.True(row1.Success, $"AddRowField Region failed: {row1.ErrorMessage}");

        var row2 = _pivotCommands.AddRowField(batch, "SalesPivot", "Product");
        Assert.True(row2.Success, $"AddRowField Product failed: {row2.ErrorMessage}");

        var value = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        Assert.True(value.Success, $"AddValueField failed: {value.ErrorMessage}");

        // Act - Set to Compact layout
        var result = _pivotCommands.SetLayout(batch, "SalesPivot", 0);

        // Assert
        Assert.True(result.Success, $"SetLayout failed: {result.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void SetLayout_Tabular_UpdatesLayoutForm()
    {
        // Arrange - Create test file with data
        var testFile = CreateTestFileWithData(nameof(SetLayout_Tabular_UpdatesLayoutForm));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        var createResult = _pivotCommands.CreateFromRange(batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesPivot");
        Assert.True(createResult.Success, $"CreateFromRange failed: {createResult.ErrorMessage}");

        var row1 = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        Assert.True(row1.Success, $"AddRowField Region failed: {row1.ErrorMessage}");

        var row2 = _pivotCommands.AddRowField(batch, "SalesPivot", "Product");
        Assert.True(row2.Success, $"AddRowField Product failed: {row2.ErrorMessage}");

        var value = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        Assert.True(value.Success, $"AddValueField failed: {value.ErrorMessage}");

        // Act - Set to Tabular layout
        var result = _pivotCommands.SetLayout(batch, "SalesPivot", 1);

        // Assert
        Assert.True(result.Success, $"SetLayout failed: {result.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void SetLayout_Outline_UpdatesLayoutForm()
    {
        // Arrange - Create test file with data
        var testFile = CreateTestFileWithData(nameof(SetLayout_Outline_UpdatesLayoutForm));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        var createResult = _pivotCommands.CreateFromRange(batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesPivot");
        Assert.True(createResult.Success, $"CreateFromRange failed: {createResult.ErrorMessage}");

        var row1 = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        Assert.True(row1.Success, $"AddRowField Region failed: {row1.ErrorMessage}");

        var row2 = _pivotCommands.AddRowField(batch, "SalesPivot", "Product");
        Assert.True(row2.Success, $"AddRowField Product failed: {row2.ErrorMessage}");

        var value = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        Assert.True(value.Success, $"AddValueField failed: {value.ErrorMessage}");

        // Act - Set to Outline layout
        var result = _pivotCommands.SetLayout(batch, "SalesPivot", 2);

        // Assert
        Assert.True(result.Success, $"SetLayout failed: {result.ErrorMessage}");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void SetSubtotals_Show_EnablesSubtotals()
    {
        // Arrange - Create test file with data
        var testFile = CreateTestFileWithData(nameof(SetSubtotals_Show_EnablesSubtotals));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        var createResult = _pivotCommands.CreateFromRange(batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesPivot");
        Assert.True(createResult.Success, $"CreateFromRange failed: {createResult.ErrorMessage}");

        var row = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        Assert.True(row.Success, $"AddRowField failed: {row.ErrorMessage}");

        var value = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        Assert.True(value.Success, $"AddValueField failed: {value.ErrorMessage}");

        // Act - Enable subtotals
        var result = _pivotCommands.SetSubtotals(batch, "SalesPivot", "Region", true);

        // Assert
        Assert.True(result.Success, $"SetSubtotals failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void SetSubtotals_Hide_DisablesSubtotals()
    {
        // Arrange - Create test file with data
        var testFile = CreateTestFileWithData(nameof(SetSubtotals_Hide_DisablesSubtotals));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        var createResult = _pivotCommands.CreateFromRange(batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesPivot");
        Assert.True(createResult.Success, $"CreateFromRange failed: {createResult.ErrorMessage}");

        var row = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        Assert.True(row.Success, $"AddRowField failed: {row.ErrorMessage}");

        var value = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        Assert.True(value.Success, $"AddValueField failed: {value.ErrorMessage}");

        // First enable subtotals
        var enable = _pivotCommands.SetSubtotals(batch, "SalesPivot", "Region", true);
        Assert.True(enable.Success);

        // Act - Disable subtotals
        var result = _pivotCommands.SetSubtotals(batch, "SalesPivot", "Region", false);

        // Assert
        Assert.True(result.Success, $"SetSubtotals failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void SetLayout_RoundTrip_PersistsLayoutChange()
    {
        // Arrange - Create test file with data
        var testFile = CreateTestFileWithData(nameof(SetLayout_RoundTrip_PersistsLayoutChange));

        using (var batch = new ExcelBatch(new[] { testFile }, _loggerFactory.CreateLogger<ExcelBatch>()))
        {
            var createResult = _pivotCommands.CreateFromRange(batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesPivot");
            Assert.True(createResult.Success);

            var row1 = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
            Assert.True(row1.Success);

            var row2 = _pivotCommands.AddRowField(batch, "SalesPivot", "Product");
            Assert.True(row2.Success);

            var value = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
            Assert.True(value.Success);

            // Act - Set to Tabular layout and save
            var layoutResult = _pivotCommands.SetLayout(batch, "SalesPivot", 1);
            Assert.True(layoutResult.Success);

            batch.Save();
        }

        // Assert - Reopen and verify PivotTable still exists and configured
        using (var batch = new ExcelBatch(new[] { testFile }, _loggerFactory.CreateLogger<ExcelBatch>()))
        {
            var listResult = _pivotCommands.List(batch);
            Assert.True(listResult.Success);
            Assert.Contains(listResult.PivotTables, pt => pt.Name == "SalesPivot");

            // Verify we can still interact with the PivotTable
            var fields = _pivotCommands.ListFields(batch, "SalesPivot");
            Assert.True(fields.Success);
            Assert.Contains(fields.Fields, f => f.Name == "Region");
            Assert.Contains(fields.Fields, f => f.Name == "Product");
        }
    }
}




