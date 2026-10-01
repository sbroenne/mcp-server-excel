using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for PivotTable field operations (Strategy Pattern: RegularPivotTableFieldStrategy).
/// Tests AddColumn, AddValue, AddFilter, Remove, Set* operations on Regular PivotTables.
/// Optimized: Single batch per test, no SaveAsync() unless testing persistence.
/// Organized by category trait for Architecture Pattern clarity.
/// </summary>
public sealed partial class PersistentServicePivotTableTests
{
    /// <inheritdoc/>
    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public void AddColumnField_WithValidField_AddsFieldToColumns()
    {
        // Arrange

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        _commands.SetValues(
            batch,
            _salesSheetName,
            "A1:D11",
            [
                ["Date", "Region", "Product", "Amount"],
                [new DateTime(2026, 1, 1), "North", "Widget", 100],
                [new DateTime(2026, 1, 2), "South", "Widget", 150],
                [new DateTime(2026, 1, 3), "North", "Gadget", 200],
                [new DateTime(2026, 1, 4), "West", "Widget", 75],
                [new DateTime(2026, 1, 5), "East", "Gadget", 125],
                [new DateTime(2026, 1, 6), "North", "Widget", 90],
                [new DateTime(2026, 1, 7), "South", "Gadget", 60],
                [new DateTime(2026, 1, 8), "West", "Widget", 110],
                [new DateTime(2026, 1, 9), "East", "Gadget", 80],
                [new DateTime(2026, 1, 10), "North", "Gadget", 130],
            ],
            overwritePolicy: OverwritePolicy.Allow);
        _commands.SetNumberFormat(batch, _salesSheetName, "A2:A11", "m/d/yyyy");
        _commands.SetNumberFormat(batch, _salesSheetName, "D2:D11", "$#,##0.00");
        _tableCommands.Create(
            batch,
            _salesSheetName,
            "tblSales",
            "A1:D11",
            true,
            TableStylePresets.Medium2);

        var createResult = _pivotCommands.CreateFromTable(
            batch, "tblSales", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

