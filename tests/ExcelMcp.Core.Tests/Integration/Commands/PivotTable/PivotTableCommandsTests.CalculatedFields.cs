using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for PivotTable calculated field operations.
/// Calculated fields create custom fields with formulas for analysis.
/// Regular PivotTables: Full support via CalculatedFields.Add() API.
/// OLAP PivotTables: NOT supported (use DAX measures instead).
/// </summary>
public partial class PivotTableCommandsTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateCalculatedField_MultiplicationFormula_CreatesField()
    {
        // Arrange - Test data has: Region, Product, Sales, Date
        var testFile = CreateTestFileWithData(nameof(CreateCalculatedField_MultiplicationFormula_CreatesField));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(testFile, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesPivot");
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
        var testFile = CreateTestFileWithData(nameof(CreateCalculatedField_SubtractionFormula_CreatesField));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(testFile, logger);

        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesPivot");
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
        var testFile = CreateTestFileWithData(nameof(CreateCalculatedField_ComplexFormula_CreatesField));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(testFile, logger);

        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesPivot");
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
        var testFile = CreateTestFileWithData(nameof(CreateCalculatedField_AdditionFormula_CreatesField));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(testFile, logger);

        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "SalesPivot");
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
}
