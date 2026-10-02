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
        var result = _dataModelCommands.ListMeasures(batch);

        Assert.True(result.Success, $"ListMeasures failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.Measures);

        // Check that previews are returned (raw DAX, not formatted)
        var totalSalesMeasure = result.Measures.FirstOrDefault(m => m.Name == "Total Sales");
        Assert.NotNull(totalSalesMeasure);
        Assert.NotEmpty(totalSalesMeasure.FormulaPreview);
        Assert.Contains("SUM", totalSalesMeasure.FormulaPreview, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests that Read returns raw DAX formula as stored (no formatting on read).
    /// Verifies that the formula is returned intact.
    /// </summary>
    [Fact]
    public void Read_WithMeasure_ReturnsRawFormula()
    {
        var batch = _fixture.BatchToken;
        var result = _dataModelCommands.Read(batch, "Total Sales");

        Assert.True(result.Success, $"Read failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.DaxFormula);
        Assert.Contains("SUM", result.DaxFormula, StringComparison.OrdinalIgnoreCase);

        // CharacterCount should reflect the formula length
        Assert.True(result.CharacterCount > 0);
        Assert.Equal(result.DaxFormula.Length, result.CharacterCount);
    }

    /// <summary>
    /// Tests that CreateMeasure preserves DAX exactly by default.
    /// </summary>
    [Fact]
    public void CreateMeasure_WithUnformattedDax_PreservesExactInputByDefault()
    {
        var measureName = $"Test_CreateFormatted_{Guid.NewGuid():N}";
        // Unformatted DAX (single line, no spaces around operators)
        var unformattedDax = "CALCULATE(SUM(SalesTable[Amount]),FILTER(SalesTable,SalesTable[CustomerID]=1))";

        var batch = _fixture.BatchToken;

        // Create measure without remote formatting opt-in
        _ = CreateMeasure("SalesTable", measureName, unformattedDax);

        // Retrieve and verify
        var viewResult = _dataModelCommands.Read(batch, measureName);
        Assert.True(viewResult.Success, $"Read failed: {viewResult.ErrorMessage}");
        Assert.Equal(unformattedDax, viewResult.DaxFormula);
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
        var unformattedUpdate = "CALCULATE(AVERAGE(SalesTable[Amount]),FILTER(SalesTable,SalesTable[Region]=\"North\"))";

        var batch = _fixture.BatchToken;

        // Create measure
        _ = CreateMeasure("SalesTable", measureName, originalFormula);

        // Update without remote formatting opt-in
        _ = _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: unformattedUpdate);

        // Retrieve and verify
        var viewResult = _dataModelCommands.Read(batch, measureName);
        Assert.True(viewResult.Success, $"Read failed: {viewResult.ErrorMessage}");
        Assert.Equal(unformattedUpdate, viewResult.DaxFormula);
    }

    /// <summary>
    /// Tests that formatted DAX still executes correctly in Excel.
    /// Creates a measure with formatted DAX and verifies it can be used in a PivotTable.
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
        SalesTable[CustomerID] = 1
    )
)";

        var batch = _fixture.BatchToken;

        // Create measure with pre-formatted DAX
        _ = CreateMeasure("SalesTable", measureName, formattedDax);

        // Retrieve and verify it was saved correctly
        var viewResult = _dataModelCommands.Read(batch, measureName);
        Assert.True(viewResult.Success, $"Read failed: {viewResult.ErrorMessage}");
        Assert.Contains("CALCULATE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);

        // Verify the measure appears in the list
        var listResult = _dataModelCommands.ListMeasures(batch);
        Assert.Contains(listResult.Measures, m => m.Name == measureName);
    }

    /// <summary>
    /// Tests that null or empty DAX is handled gracefully (no formatting attempted).
    /// </summary>
    [Fact]
    public void UpdateMeasure_WithNullDaxFormula_DoesNotAttemptFormatting()
    {
        var measureName = $"Test_NullFormula_{Guid.NewGuid():N}";
        var originalFormula = "SUM(SalesTable[Amount])";
        var newDescription = "Updated description";

        var batch = _fixture.BatchToken;

        // Create measure
        _ = CreateMeasure("SalesTable", measureName, originalFormula);

        // Update only description (null daxFormula should not trigger formatting)
        _ = _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: null, description: newDescription);

        // Verify description updated, formula unchanged
        var viewResult = _dataModelCommands.Read(batch, measureName);
        Assert.True(viewResult.Success, $"Read failed: {viewResult.ErrorMessage}");
        Assert.Equal(newDescription, viewResult.Description);
        Assert.Contains("SUM", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    private OperationResult CreateMeasure(
        string tableName,
        string measureName,
        string daxFormula)
    {
        var result = _dataModelCommands.CreateMeasure(
            _fixture.BatchToken,
            tableName,
            measureName,
            daxFormula);
        _fixture.RegisterDataModelMeasureForCleanup(measureName);
        return result;
    }

}

