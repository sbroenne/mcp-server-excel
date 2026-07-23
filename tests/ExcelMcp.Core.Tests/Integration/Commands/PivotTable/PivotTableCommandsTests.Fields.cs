using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for PivotTable field operations (Strategy Pattern: RegularPivotTableFieldStrategy).
/// Tests AddColumn, AddValue, AddFilter, Remove, Set* operations on Regular PivotTables.
/// Optimized: Single batch per test, no SaveAsync() unless testing persistence.
/// Organized by category trait for Architecture Pattern clarity.
/// </summary>
public partial class PivotTableCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void AddColumnField_WithValidField_AddsFieldToColumns()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(AddColumnField_WithValidField_AddsFieldToColumns));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act - No save needed
        var result = _pivotCommands.AddColumnField(batch, "TestPivot", "Product");

        // Assert
        Assert.True(result.Success, $"AddColumnField failed: {result.ErrorMessage}");
        Assert.Equal("Product", result.FieldName);
        Assert.Equal(PivotFieldArea.Column, result.Area);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void AddValueField_WithValidField_AddsFieldToValues()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(AddValueField_WithValidField_AddsFieldToValues));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act
        var result = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");

        // Assert
        Assert.True(result.Success, $"AddValueField failed: {result.ErrorMessage}");
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal(PivotFieldArea.Value, result.Area);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void AddValueField_FromTableNumericColumn_AllowsSumAggregation()
    {
        // Arrange
        var testFile = _olapFixture.CreateTestFile(nameof(AddValueField_FromTableNumericColumn_AllowsSumAggregation));

        using var batch = ExcelSession.BeginBatch(testFile);
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ctx.Book.Worksheets[1];
                sheet.Name = "SalesData";

                sheet.Range["A1"].Value2 = "Date";
                sheet.Range["B1"].Value2 = "Region";
                sheet.Range["C1"].Value2 = "Product";
                sheet.Range["D1"].Value2 = "Amount";

                sheet.Range["A2"].Value2 = new DateTime(2026, 1, 1);
                sheet.Range["B2"].Value2 = "North";
                sheet.Range["C2"].Value2 = "Widget";
                sheet.Range["D2"].Value2 = 100;

                sheet.Range["A3"].Value2 = new DateTime(2026, 1, 2);
                sheet.Range["B3"].Value2 = "South";
                sheet.Range["C3"].Value2 = "Widget";
                sheet.Range["D3"].Value2 = 150;

                sheet.Range["A4"].Value2 = new DateTime(2026, 1, 3);
                sheet.Range["B4"].Value2 = "North";
                sheet.Range["C4"].Value2 = "Gadget";
                sheet.Range["D4"].Value2 = 200;

                sheet.Range["A5"].Value2 = new DateTime(2026, 1, 4);
                sheet.Range["B5"].Value2 = "West";
                sheet.Range["C5"].Value2 = "Widget";
                sheet.Range["D5"].Value2 = 75;

                sheet.Range["A6"].Value2 = new DateTime(2026, 1, 5);
                sheet.Range["B6"].Value2 = "East";
                sheet.Range["C6"].Value2 = "Gadget";
                sheet.Range["D6"].Value2 = 125;

                sheet.Range["A7"].Value2 = new DateTime(2026, 1, 6);
                sheet.Range["B7"].Value2 = "North";
                sheet.Range["C7"].Value2 = "Widget";
                sheet.Range["D7"].Value2 = 90;

                sheet.Range["A8"].Value2 = new DateTime(2026, 1, 7);
                sheet.Range["B8"].Value2 = "South";
                sheet.Range["C8"].Value2 = "Gadget";
                sheet.Range["D8"].Value2 = 60;

                sheet.Range["A9"].Value2 = new DateTime(2026, 1, 8);
                sheet.Range["B9"].Value2 = "West";
                sheet.Range["C9"].Value2 = "Widget";
                sheet.Range["D9"].Value2 = 110;

                sheet.Range["A10"].Value2 = new DateTime(2026, 1, 9);
                sheet.Range["B10"].Value2 = "East";
                sheet.Range["C10"].Value2 = "Gadget";
                sheet.Range["D10"].Value2 = 80;

                sheet.Range["A11"].Value2 = new DateTime(2026, 1, 10);
                sheet.Range["B11"].Value2 = "North";
                sheet.Range["C11"].Value2 = "Gadget";
                sheet.Range["D11"].Value2 = 130;

                sheet.Range["A2:A11"].NumberFormat = "m/d/yyyy";
                sheet.Range["D2:D11"].NumberFormat = "$#,##0.00";

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });

        var tableCommands = new TableCommands();
        tableCommands.Create(batch, "SalesData", "tblSales", "A1:D11", true, TableStylePresets.Medium2);

        var createResult = _pivotCommands.CreateFromTable(
            batch, "tblSales", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success, $"CreateFromTable failed: {createResult.ErrorMessage}");

        // Act
        var fieldsResult = _pivotCommands.ListFields(batch, "TestPivot");
        var addResult = _pivotCommands.AddValueField(
            batch, "TestPivot", "Amount", AggregationFunction.Sum, "Total Amount");

        // Assert
        Assert.True(fieldsResult.Success, $"ListFields failed: {fieldsResult.ErrorMessage}");
        var amountField = Assert.Single(fieldsResult.Fields, field => field.Name == "Amount");
        Assert.Equal("Number", amountField.DataType);

        Assert.True(addResult.Success, $"AddValueField failed: {addResult.ErrorMessage}");
        Assert.Equal("Amount", addResult.FieldName);
        Assert.Equal(PivotFieldArea.Value, addResult.Area);
        Assert.Equal(AggregationFunction.Sum, addResult.Function);
        Assert.Equal("Number", addResult.DataType);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void AddFilterField_WithValidField_AddsFieldToFilters()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(AddFilterField_WithValidField_AddsFieldToFilters));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act
        var result = _pivotCommands.AddFilterField(batch, "TestPivot", "Region");

        // Assert
        Assert.True(result.Success, $"AddFilterField failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(PivotFieldArea.Filter, result.Area);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void RemoveField_ExistingField_RemovesFromPivot()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(RemoveField_ExistingField_RemovesFromPivot));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add a field first
        var addResult = _pivotCommands.AddRowField(batch, "TestPivot", "Region");
        Assert.True(addResult.Success);

        // Act - Remove in same batch
        var result = _pivotCommands.RemoveField(batch, "TestPivot", "Region");

        // Assert
        Assert.True(result.Success, $"RemoveField failed: {result.ErrorMessage}");

        // Verify field removed
        var infoResult = _pivotCommands.Read(batch, "TestPivot");
        Assert.True(infoResult.Success);
        var regionField = infoResult.Fields.FirstOrDefault(f => f.Name == "Region");
        Assert.NotNull(regionField);
        Assert.Equal(PivotFieldArea.Hidden, regionField.Area);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldFunction_ValueField_ChangesAggregation()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetFieldFunction_ValueField_ChangesAggregation));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add Sales as value field (default sum)
        var addResult = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");
        Assert.True(addResult.Success);

        // Act - Change to Average in same batch
        var result = _pivotCommands.SetFieldFunction(batch, "TestPivot", "Sales", AggregationFunction.Average);

        // Assert
        Assert.True(result.Success, $"SetFieldFunction failed: {result.ErrorMessage}");
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal(AggregationFunction.Average, result.Function);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldName_ExistingField_RenamesField()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetFieldName_ExistingField_RenamesField));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add Sales as value field
        var addResult = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");
        Assert.True(addResult.Success);

        // Act
        var result = _pivotCommands.SetFieldName(batch, "TestPivot", "Sales", "Total Revenue");

        // Assert
        Assert.True(result.Success, $"SetFieldName failed: {result.ErrorMessage}");
        Assert.Equal("Total Revenue", result.CustomName);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldFormat_ValueField_AppliesNumberFormat()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetFieldFormat_ValueField_AppliesNumberFormat));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add Sales as value field
        var addResult = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");
        Assert.True(addResult.Success);

        // Act
        var result = _pivotCommands.SetFieldFormat(batch, "TestPivot", "Sales", "$#,##0.00");

        // Assert
        Assert.True(result.Success, $"SetFieldFormat failed: {result.ErrorMessage}");
        // Note: Excel COM may normalize format codes. We just verify a format was applied and contains currency/decimal indicators
        Assert.NotNull(result.NumberFormat);
        Assert.Contains("$", result.NumberFormat);
    }

    /// <summary>
    /// Verifies that SetFieldFormat with US currency format works correctly on any locale.
    /// The server should auto-translate format codes (number separators handled by UseSystemSeparators=false).
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldFormat_USCurrencyFormat_RoundTripsCorrectly()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetFieldFormat_USCurrencyFormat_RoundTripsCorrectly));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add Sales as value field
        var addResult = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");
        Assert.True(addResult.Success);

        // Act - Apply US currency format
        var result = _pivotCommands.SetFieldFormat(batch, "TestPivot", "Sales", "$#,##0.00");

        // Assert - Format should round-trip correctly (not corrupted by locale)
        Assert.True(result.Success, $"SetFieldFormat failed: {result.ErrorMessage}");
        Assert.NotNull(result.NumberFormat);
        // Verify the format contains expected components ($ symbol, decimal separator)
        Assert.Contains("$", result.NumberFormat);
        Assert.Contains(".", result.NumberFormat); // Decimal separator should be preserved
        Assert.Contains(",", result.NumberFormat); // Thousands separator should be preserved
    }

    /// <summary>
    /// Verifies that SetFieldFormat with US date format works correctly on value fields.
    /// Uses a Count function on a date field to create a numeric value that can be formatted.
    /// The server auto-translates format codes (number separators handled by UseSystemSeparators=false).
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldFormat_USPercentFormat_RoundTripsCorrectly()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetFieldFormat_USPercentFormat_RoundTripsCorrectly));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add Sales as value field
        var addResult = _pivotCommands.AddValueField(batch, "TestPivot", "Sales");
        Assert.True(addResult.Success);

        // Act - Apply US percent format (tests decimal separator preservation)
        var result = _pivotCommands.SetFieldFormat(batch, "TestPivot", "Sales", "0.00%");

        // Assert - Format should round-trip correctly (not corrupted by locale)
        Assert.True(result.Success, $"SetFieldFormat failed: {result.ErrorMessage}");
        Assert.NotNull(result.NumberFormat);
        // Verify the format contains expected components (percent symbol, decimal separator)
        Assert.Contains("%", result.NumberFormat);
        Assert.Contains(".", result.NumberFormat); // Decimal separator should be preserved
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SetFieldFilter_RowField_AppliesFilter()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetFieldFilter_RowField_AppliesFilter));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add Region as row field
        var addResult = _pivotCommands.AddRowField(batch, "TestPivot", "Region");
        Assert.True(addResult.Success);

        // Act
        var result = _pivotCommands.SetFieldFilter(batch, "TestPivot", "Region", ["North"]);

        // Assert
        Assert.True(result.Success, $"SetFieldFilter failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.NotEmpty(result.SelectedItems);
    }
    /// <inheritdoc/>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void SortField_RowField_SortsData()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SortField_RowField_SortsData));

        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add Region as row field
        var addResult = _pivotCommands.AddRowField(batch, "TestPivot", "Region");
        Assert.True(addResult.Success);

        // Act
        var result = _pivotCommands.SortField(batch, "TestPivot", "Region", SortDirection.Ascending);

        // Assert
        Assert.True(result.Success, $"SortField failed: {result.ErrorMessage}");
    }
}




