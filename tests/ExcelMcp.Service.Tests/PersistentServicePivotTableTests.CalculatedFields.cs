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
        RequireSuccess(created);
        var row = _pivotCommands.AddRowField(batch, "CalculatedTotals", "Region");
        RequireSuccess(row);
        var calculated = _pivotCommands.CreateCalculatedField(
            batch, "CalculatedTotals", "DoubleSales", "=Sales*2");
        RequireSuccess(calculated);

        var added = _pivotCommands.AddValueField(batch, "CalculatedTotals", "DoubleSales");
        RequireSuccess(added);
        Assert.Equal("Number", added.DataType);
        var configured = _pivotCommands.SetFieldFunction(
            batch, "CalculatedTotals", "DoubleSales", AggregationFunction.Sum);
        RequireSuccess(configured);
        var refreshed = _pivotCommands.Refresh(batch, "CalculatedTotals");
        RequireSuccess(refreshed);

        var data = _pivotCommands.GetData(batch, "CalculatedTotals");
        RequireSuccess(data);
        Assert.Equal(1300d, Convert.ToDouble(data.Values[^1][^1], CultureInfo.InvariantCulture));
        var fields = _pivotCommands.ListFields(batch, "CalculatedTotals");
        RequireSuccess(fields);
        Assert.Equal("Number", Assert.Single(fields.Fields, field => field.Name == "DoubleSales").DataType);
        AssertPivotSales(650, 650, "CalculatedTotals");
        AssertOriginalSales();
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
        RequireSuccess(createResult);

        // Add fields
        var rowResult = _pivotCommands.AddRowField(batch, "SalesPivot", "Product");
        RequireSuccess(rowResult);

        var valueResult = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        RequireSuccess(valueResult);

        // Act - Create calculated field (Sales * 2)
        var result = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "DoubleSales", "=Sales*2");

        // Assert
        RequireSuccess(result);
        Assert.Equal("DoubleSales", result.FieldName);
        Assert.Equal("=Sales*2", result.Formula);
        Assert.NotNull(result.WorkflowHint);
        Assert.Contains("Add to Values area", result.WorkflowHint);

        // Verify field exists
        var listResult = _pivotCommands.ListFields(batch, "SalesPivot");
        RequireSuccess(listResult);
        Assert.Contains(listResult.Fields, f => f.Name == "DoubleSales");
        AssertCalculatedValues("SalesPivot", "DoubleSales",
            new Dictionary<string, double> { ["Gadget"] = 550, ["Widget"] = 750 }, 1300);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateCalculatedField_SubtractionFormula_CreatesField()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        RequireSuccess(createResult);

        var rowResult = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        RequireSuccess(rowResult);

        var valueResult = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        RequireSuccess(valueResult);

        // Act - Subtraction formula (Sales - 100)
        var result = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "AfterDiscount", "=Sales-100");

        // Assert
        RequireSuccess(result);
        Assert.Equal("AfterDiscount", result.FieldName);
        Assert.Equal("=Sales-100", result.Formula);

        var listResult = _pivotCommands.ListFields(batch, "SalesPivot");
        RequireSuccess(listResult);
        Assert.Contains(listResult.Fields, f => f.Name == "AfterDiscount");
        AssertCalculatedValues("SalesPivot", "AfterDiscount",
            new Dictionary<string, double> { ["North"] = 225, ["South"] = 225 }, 550);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateCalculatedField_ComplexFormula_CreatesField()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        RequireSuccess(createResult);

        var rowResult = _pivotCommands.AddRowField(batch, "SalesPivot", "Product");
        RequireSuccess(rowResult);

        var valueResult = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        RequireSuccess(valueResult);

        // Act - Complex formula with parentheses: (Sales - 50) / Sales
        var result = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "ProfitMargin", "=(Sales-50)/Sales");

        // Assert
        RequireSuccess(result);
        Assert.Equal("ProfitMargin", result.FieldName);
        Assert.Equal("=(Sales-50)/Sales", result.Formula);

        var listResult = _pivotCommands.ListFields(batch, "SalesPivot");
        RequireSuccess(listResult);
        Assert.Contains(listResult.Fields, f => f.Name == "ProfitMargin");
        AssertCalculatedValues("SalesPivot", "ProfitMargin",
            new Dictionary<string, double> { ["Gadget"] = 225d / 275, ["Widget"] = 325d / 375 }, 600d / 650);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateCalculatedField_AdditionFormula_CreatesField()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        RequireSuccess(createResult);

        var rowResult = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        RequireSuccess(rowResult);

        var valueResult = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        RequireSuccess(valueResult);

        // Act - Addition formula (Sales + 50 as bonus)
        var result = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "WithBonus", "=Sales+50");

        // Assert
        RequireSuccess(result);
        Assert.Equal("WithBonus", result.FieldName);
        Assert.Equal("=Sales+50", result.Formula);

        var listResult = _pivotCommands.ListFields(batch, "SalesPivot");
        RequireSuccess(listResult);
        Assert.Contains(listResult.Fields, f => f.Name == "WithBonus");
        AssertCalculatedValues("SalesPivot", "WithBonus",
            new Dictionary<string, double> { ["North"] = 375, ["South"] = 375 }, 700);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void ListCalculatedFields_NoCalculatedFields_ReturnsEmptyList()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        RequireSuccess(createResult);

        // Act - List calculated fields (should be empty)
        var result = _pivotCommands.ListCalculatedFields(batch, "SalesPivot");

        // Assert
        RequireSuccess(result);
        Assert.Empty(result.CalculatedFields);
        RequireSuccess(result);
        AssertOriginalSales();
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void ListCalculatedFields_AfterCreate_ReturnsField()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createPivotResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        RequireSuccess(createPivotResult);

        // Add a row field and value to make valid PivotTable
        RequireSuccess(_pivotCommands.AddRowField(batch, "SalesPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "SalesPivot", "Sales"));

        // Create a calculated field
        var createResult = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "DoubleSales", "=Sales*2");
        RequireSuccess(createResult);

        // Act - List calculated fields
        var result = _pivotCommands.ListCalculatedFields(batch, "SalesPivot");

        // Assert
        RequireSuccess(result);
        Assert.Single(result.CalculatedFields);
        Assert.Equal("DoubleSales", result.CalculatedFields[0].Name);
        RequireSuccess(result);
        Assert.Equal("=Sales*2", result.CalculatedFields[0].Formula);
        AssertOriginalSales();
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteCalculatedField_ExistingField_RemovesField()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createPivotResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        RequireSuccess(createPivotResult);

        RequireSuccess(_pivotCommands.AddRowField(batch, "SalesPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "SalesPivot", "Sales"));

        // Create a calculated field
        var createResult = _pivotCommands.CreateCalculatedField(batch, "SalesPivot", "TestCalcField", "=Sales*3");
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "SalesPivot", "Retained", "=Sales*2"));

        // Verify it exists
        var listBefore = _pivotCommands.ListCalculatedFields(batch, "SalesPivot");
        Assert.Contains(listBefore.CalculatedFields, f => f.Name == "TestCalcField");
        RequireSuccess(listBefore);

        // Act - Delete the calculated field
        var deleteResult = _pivotCommands.DeleteCalculatedField(batch, "SalesPivot", "TestCalcField");

        // Assert
        RequireSuccess(deleteResult);

        // Verify it's gone
        var listAfter = _pivotCommands.ListCalculatedFields(batch, "SalesPivot");
        Assert.DoesNotContain(listAfter.CalculatedFields, f => f.Name == "TestCalcField");
        RequireSuccess(deleteResult);
        RequireSuccess(listAfter);
        var retained = Assert.Single(listAfter.CalculatedFields);
        Assert.Equal("Retained", retained.Name);
        Assert.Equal("=Sales*2", retained.Formula);
        AssertOriginalSales();
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteCalculatedField_NonExistentField_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        var createPivotResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SalesPivot");
        RequireSuccess(createPivotResult);
        RequireSuccess(_pivotCommands.CreateCalculatedField(batch, "SalesPivot", "Retained", "=Sales*2"));
        var before = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_pivotCommands.ListCalculatedFields(batch, "SalesPivot")).CalculatedFields);

        // Act - Try to delete non-existent field
        var result = _pivotCommands.DeleteCalculatedField(batch, "SalesPivot", "NonExistentField");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage);
        Assert.Equal(before, System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_pivotCommands.ListCalculatedFields(batch, "SalesPivot")).CalculatedFields));
        AssertOriginalSales();
    }

    private void AssertCalculatedValues(
        string pivotName, string fieldName, Dictionary<string, double> expected, double total)
    {
        RequireSuccess(_pivotCommands.AddValueField(_fixture.BatchToken, pivotName, fieldName));
        RequireSuccess(_pivotCommands.Refresh(_fixture.BatchToken, pivotName));
        var values = RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, pivotName)).Values;
        Assert.Equal(expected.Count + 2, values.Count);
        Assert.All(values, row => Assert.Equal(3, row.Count));
        foreach (var pair in expected)
        {
            var row = Assert.Single(values, row => string.Equals(row[0]?.ToString(), pair.Key, StringComparison.Ordinal));
            Assert.Equal(pair.Value, Convert.ToDouble(row[^1], CultureInfo.InvariantCulture), 10);
        }
        Assert.Equal(total, Convert.ToDouble(values[^1][^1], CultureInfo.InvariantCulture), 10);
        AssertOriginalSales();
    }
}

