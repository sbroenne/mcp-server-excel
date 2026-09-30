using System.Globalization;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for PivotTable calculated field operations.
/// Calculated fields create custom fields with formulas for analysis.
/// Regular PivotTables: Full support via CalculatedFields.Add() API.
/// OLAP PivotTables: NOT supported (use DAX measures instead).
/// </summary>
public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public void CalculatedField_AddAndSetSum_ReturnsNumericTotals()
    {
        var batch = _fixture.BatchToken;
        var created = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "CalculatedTotals");
        Assert.True(created.Success, created.ErrorMessage);
        var row = _pivotCommands.AddRowField(batch, "CalculatedTotals", "Region");
        Assert.True(row.Success, row.ErrorMessage);
        var calculated = _pivotCommands.CreateCalculatedField(
            batch, "CalculatedTotals", "DoubleSales", "=Sales*2");
        Assert.True(calculated.Success, calculated.ErrorMessage);

        var added = _pivotCommands.AddValueField(batch, "CalculatedTotals", "DoubleSales");
        Assert.True(added.Success, added.ErrorMessage);
        Assert.Equal("Number", added.DataType);
        var configured = _pivotCommands.SetFieldFunction(
            batch, "CalculatedTotals", "DoubleSales", AggregationFunction.Sum);
        Assert.True(configured.Success, configured.ErrorMessage);
        var refreshed = _pivotCommands.Refresh(batch, "CalculatedTotals");
        Assert.True(refreshed.Success, refreshed.ErrorMessage);

        var data = _pivotCommands.GetData(batch, "CalculatedTotals");
        Assert.True(data.Success, data.ErrorMessage);
        Assert.Equal(1300d, Convert.ToDouble(data.Values[^1][^1], CultureInfo.InvariantCulture));
        var fields = _pivotCommands.ListFields(batch, "CalculatedTotals");
        Assert.True(fields.Success, fields.ErrorMessage);
        Assert.Equal("Number", Assert.Single(fields.Fields, field => field.Name == "DoubleSales").DataType);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateCalculatedField_MultiplicationFormula_CreatesField()
    {
        // Arrange - Test data has: Region, Product, Sales, Date
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        Assert.True(createResult.Success, $"CreateFromRange failed: {createResult.ErrorMessage}");

        // Add fields
        var rowResult = _pivotCommands.AddRowField(batch, "SalesPivot", "Product");
        Assert.True(rowResult.Success, $"AddRowField failed: {rowResult.ErrorMessage}");

        var valueResult = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        Assert.True(valueResult.Success, $"AddValueField failed: {valueResult.ErrorMessage}");

        // Act - Create calculated field (Sales * 2)
        var result = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "DoubleSales", "=Sales*2");

        // Assert
        Assert.True(result.Success, $"CreateCalculatedField failed: {result.ErrorMessage}");
        Assert.Equal("DoubleSales", result.FieldName);
        Assert.Equal("=Sales*2", result.Formula);
        Assert.NotNull(result.WorkflowHint);
        Assert.Contains("Add to Values area", result.WorkflowHint);

        // Verify field exists
        var listResult = _pivotCommands.ListFields(batch, "SalesPivot");
        Assert.True(listResult.Success, $"ListFields failed: {listResult.ErrorMessage}");
        Assert.Contains(listResult.Fields, f => f.Name == "DoubleSales");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateCalculatedField_SubtractionFormula_CreatesField()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        Assert.True(createResult.Success);

        var rowResult = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        Assert.True(rowResult.Success);

        var valueResult = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        Assert.True(valueResult.Success);

        // Act - Subtraction formula (Sales - 100)
        var result = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "AfterDiscount", "=Sales-100");

        // Assert
        Assert.True(result.Success, $"CreateCalculatedField failed: {result.ErrorMessage}");
        Assert.Equal("AfterDiscount", result.FieldName);
        Assert.Equal("=Sales-100", result.Formula);

        var listResult = _pivotCommands.ListFields(batch, "SalesPivot");
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Fields, f => f.Name == "AfterDiscount");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateCalculatedField_ComplexFormula_CreatesField()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        Assert.True(createResult.Success);

        var rowResult = _pivotCommands.AddRowField(batch, "SalesPivot", "Product");
        Assert.True(rowResult.Success);

        var valueResult = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        Assert.True(valueResult.Success);

        // Act - Complex formula with parentheses: (Sales - 50) / Sales
        var result = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "ProfitMargin", "=(Sales-50)/Sales");

        // Assert
        Assert.True(result.Success, $"CreateCalculatedField failed: {result.ErrorMessage}");
        Assert.Equal("ProfitMargin", result.FieldName);
        Assert.Equal("=(Sales-50)/Sales", result.Formula);

        var listResult = _pivotCommands.ListFields(batch, "SalesPivot");
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Fields, f => f.Name == "ProfitMargin");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateCalculatedField_AdditionFormula_CreatesField()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        Assert.True(createResult.Success);

        var rowResult = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        Assert.True(rowResult.Success);

        var valueResult = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        Assert.True(valueResult.Success);

        // Act - Addition formula (Sales + 50 as bonus)
        var result = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "WithBonus", "=Sales+50");

        // Assert
        Assert.True(result.Success, $"CreateCalculatedField failed: {result.ErrorMessage}");
        Assert.Equal("WithBonus", result.FieldName);
        Assert.Equal("=Sales+50", result.Formula);

        var listResult = _pivotCommands.ListFields(batch, "SalesPivot");
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Fields, f => f.Name == "WithBonus");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void ListCalculatedFields_NoCalculatedFields_ReturnsEmptyList()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        Assert.True(createResult.Success);

        // Act - List calculated fields (should be empty)
        var result = _pivotCommands.ListCalculatedFields(batch, "SalesPivot");

        // Assert
        Assert.True(result.Success, $"ListCalculatedFields failed: {result.ErrorMessage}");
        Assert.Empty(result.CalculatedFields);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void ListCalculatedFields_AfterCreate_ReturnsField()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createPivotResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        Assert.True(createPivotResult.Success);

        // Add a row field and value to make valid PivotTable
        _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");

        // Create a calculated field
        var createResult = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "DoubleSales", "=Sales*2");
        Assert.True(createResult.Success, $"CreateCalculatedField failed: {createResult.ErrorMessage}");

        // Act - List calculated fields
        var result = _pivotCommands.ListCalculatedFields(batch, "SalesPivot");

        // Assert
        Assert.True(result.Success, $"ListCalculatedFields failed: {result.ErrorMessage}");
        Assert.Single(result.CalculatedFields);
        Assert.Equal("DoubleSales", result.CalculatedFields[0].Name);
        Assert.Contains("Sales*2", result.CalculatedFields[0].Formula);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteCalculatedField_ExistingField_RemovesField()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createPivotResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        Assert.True(createPivotResult.Success);

        _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");

        // Create a calculated field
        var createResult = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "TestCalcField", "=Sales*3");
        Assert.True(createResult.Success);

        // Verify it exists
        var listBefore = _pivotCommands.ListCalculatedFields(batch, "SalesPivot");
        Assert.Contains(listBefore.CalculatedFields, f => f.Name == "TestCalcField");

        // Act - Delete the calculated field
        var deleteResult = _pivotCommands.DeleteCalculatedField(batch, "SalesPivot", "TestCalcField");

        // Assert
        Assert.True(deleteResult.Success, $"DeleteCalculatedField failed: {deleteResult.ErrorMessage}");

        // Verify it's gone
        var listAfter = _pivotCommands.ListCalculatedFields(batch, "SalesPivot");
        Assert.DoesNotContain(listAfter.CalculatedFields, f => f.Name == "TestCalcField");
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteCalculatedField_NonExistentField_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createPivotResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        Assert.True(createPivotResult.Success);

        // Act - Try to delete non-existent field
        var result = _pivotCommands.DeleteCalculatedField(batch, "SalesPivot", "NonExistentField");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
    }
}


