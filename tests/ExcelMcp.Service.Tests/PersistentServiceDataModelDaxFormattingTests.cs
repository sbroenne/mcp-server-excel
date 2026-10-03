using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for DAX formatting behavior.
/// Tests verify that DAX formulas are preserved exactly by default.
/// Remote formatting requires explicit opt-in on CreateMeasure and UpdateMeasure.
/// </summary>
[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public class PersistentServiceDataModelDaxFormattingTests(
    PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceDataModelFixture>
{
    private readonly IDataModelCommands _dataModelCommands =
        fixture.CreateCommands<IDataModelCommands>();

    /// <summary>
    /// Tests that ListMeasures returns raw DAX previews (no formatting on read).
    /// Verifies that formula previews are returned as stored in the Data Model.
    /// </summary>
    [Fact]
    public void ListMeasures_WithMeasures_ReturnsRawPreviews()
    {
        var batch = _fixture.BatchToken;
        var result = RequireSuccess(_dataModelCommands.ListMeasures(batch));
        var totalSalesMeasure = Assert.Single(result.Measures, m => m.Name == "Total Sales");
        Assert.Equal("SalesTable", totalSalesMeasure.Table);
        Assert.Equal("SUM(SalesTable[Amount])", totalSalesMeasure.FormulaPreview);
        Assert.Equal("Total sales amount", totalSalesMeasure.Description);
        AssertMeasureValue("Total Sales", 2455);
    }

    /// <summary>
    /// Tests that Read returns raw DAX formula as stored (no formatting on read).
    /// Verifies that the formula is returned intact.
    /// </summary>
    [Fact]
    public void Read_WithMeasure_ReturnsRawFormula()
    {
        var result = AssertStoredFormula("Total Sales", "SUM(SalesTable[Amount])");
        Assert.Equal("Total sales amount", result.Description);
        Assert.Equal("Decimal", Assert.IsType<MeasureFormatInfo>(result.FormatInfo).Type);
        AssertMeasureValue("Total Sales", 2455);
    }

    /// <summary>
    /// Tests that CreateMeasure preserves DAX exactly by default.
    /// </summary>
    [Fact]
    public void CreateMeasure_WithUnformattedDax_PreservesExactInputByDefault()
    {
        var measureName = $"Test_CreateFormatted_{Guid.NewGuid():N}";
        // Unformatted DAX (single line, no spaces around operators)
        var unformattedDax = "CALCULATE(SUM(SalesTable[Amount]),FILTER(SalesTable,SalesTable[CustomerID]=101))";

        RequireSuccess(CreateMeasure("SalesTable", measureName, unformattedDax));
        AssertStoredFormula(measureName, unformattedDax);
        AssertMeasureValue(measureName, 525);
    }

    /// <summary>
    /// Tests that UpdateMeasure preserves DAX exactly by default.
    /// </summary>
    [Fact]
    public void UpdateMeasure_WithUnformattedDax_PreservesExactInputByDefault()
    {
        var measureName = $"Test_UpdateFormatted_{Guid.NewGuid():N}";
        var originalFormula = "SUM(SalesTable[Amount])";
        // Unformatted DAX for update (single line, no spaces)
        var unformattedUpdate = "CALCULATE(AVERAGE(SalesTable[Amount]),FILTER(SalesTable,RELATED(CustomersTable[Region])=\"North\"))";

        var batch = _fixture.BatchToken;

        RequireSuccess(CreateMeasure("SalesTable", measureName, originalFormula));
        AssertStoredFormula(measureName, originalFormula);
        AssertMeasureValue(measureName, 2455);
        RequireSuccess(_dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: unformattedUpdate));
        AssertStoredFormula(measureName, unformattedUpdate);
        AssertMeasureValue(measureName, 200);
    }

    /// <summary>
    /// Tests that formatted DAX still executes correctly in Excel.
    /// Creates a measure with formatted DAX and evaluates its numeric result.
    /// </summary>
    [Fact]
    public void CreateMeasure_WithFormattedDax_ExecutesCorrectlyInExcel()
    {
        var measureName = $"Test_ExecuteFormatted_{Guid.NewGuid():N}";
        // Pre-formatted DAX (with newlines and indentation)
        var formattedDax = @"CALCULATE(
    SUM(SalesTable[Amount]),
    FILTER(
        SalesTable,
        SalesTable[CustomerID] = 101
    )
)";

        RequireSuccess(CreateMeasure("SalesTable", measureName, formattedDax));
        AssertStoredFormula(measureName, formattedDax.ReplaceLineEndings("\n"));
        AssertMeasureValue(measureName, 525);
    }

    /// <summary>
    /// Tests that a description-only update preserves the formula, format, and calculation.
    /// </summary>
    [Fact]
    public void UpdateMeasure_WithNullDaxFormula_DoesNotAttemptFormatting()
    {
        var measureName = $"Test_NullFormula_{Guid.NewGuid():N}";
        var originalFormula = "SUM(SalesTable[Amount])";
        var newDescription = "Updated description";

        var batch = _fixture.BatchToken;

        RequireSuccess(CreateMeasure("SalesTable", measureName, originalFormula, formatType: "Decimal"));
        var before = AssertStoredFormula(measureName, originalFormula);
        Assert.Equal("Decimal", Assert.IsType<MeasureFormatInfo>(before.FormatInfo).Type);
        AssertMeasureValue(measureName, 2455);
        RequireSuccess(_dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: null, description: newDescription));
        var viewResult = AssertStoredFormula(measureName, originalFormula);
        Assert.Equal(newDescription, viewResult.Description);
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(before.FormatInfo),
            System.Text.Json.JsonSerializer.Serialize(viewResult.FormatInfo));
        AssertMeasureValue(measureName, 2455);
    }

    [Fact]
    public void UpdateMeasure_InvalidFormat_PreservesFormulaDescriptionFormatAndCalculation()
    {
        var measureName = $"Test_RejectedFormat_{Guid.NewGuid():N}";
        const string formula = "SUM(SalesTable[Amount])";
        RequireSuccess(CreateMeasure("SalesTable", measureName, formula,
            formatType: "Decimal", description: "Retained description"));
        var before = AssertStoredFormula(measureName, formula);
        AssertMeasureValue(measureName, 2455);
        var listed = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_dataModelCommands.ListMeasures(_fixture.BatchToken)).Measures);

        var error = Assert.Throws<ArgumentException>(() => _dataModelCommands.UpdateMeasure(
            _fixture.BatchToken, measureName, daxFormula: "0",
            formatType: "NotAFormat", description: "Rejected description"));
        Assert.Contains("format", error.Message, StringComparison.OrdinalIgnoreCase);

        var after = AssertStoredFormula(measureName, formula);
        Assert.Equal(before.Description, after.Description);
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(before.FormatInfo),
            System.Text.Json.JsonSerializer.Serialize(after.FormatInfo));
        Assert.Equal(listed, System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_dataModelCommands.ListMeasures(_fixture.BatchToken)).Measures));
        AssertMeasureValue(measureName, 2455);
    }

    private DataModelMeasureViewResult AssertStoredFormula(string measureName, string formula)
    {
        var result = RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, measureName));
        Assert.Equal(measureName, result.MeasureName);
        Assert.Equal("SalesTable", result.TableName);
        Assert.Equal(formula, result.DaxFormula);
        Assert.Equal(formula.Length, result.CharacterCount);
        var preview = Assert.Single(
            RequireSuccess(_dataModelCommands.ListMeasures(_fixture.BatchToken)).Measures,
            measure => measure.Name == measureName);
        Assert.Equal("SalesTable", preview.Table);
        Assert.Equal(formula, preview.FormulaPreview);
        Assert.Equal(result.Description, preview.Description);
        return result;
    }

    private void AssertMeasureValue(string measureName, double expected)
    {
        var result = RequireSuccess(_dataModelCommands.Evaluate(
            _fixture.BatchToken, $"EVALUATE ROW(\"Result\", [{measureName}])"));
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal("[Result]", Assert.Single(result.Columns));
        Assert.Equal(expected, Convert.ToDouble(Assert.Single(Assert.Single(result.Rows)),
            System.Globalization.CultureInfo.InvariantCulture));
    }

    private OperationResult CreateMeasure(
        string tableName,
        string measureName,
        string daxFormula,
        string? formatType = null,
        string? description = null)
    {
        var result = _dataModelCommands.CreateMeasure(
            _fixture.BatchToken,
            tableName,
            measureName,
            daxFormula,
            formatType: formatType,
            description: description);
        _fixture.RegisterDataModelMeasureForCleanup(measureName);
        return result;
    }

}
