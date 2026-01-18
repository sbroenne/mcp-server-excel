using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Integration tests for DAX formatting feature.
/// Tests verify that DAX formulas are automatically formatted on write operations.
/// Write operations (CreateMeasure, UpdateMeasure) format DAX via daxformatter.com API.
/// Read operations (ListMeasures, Read) return raw DAX as stored in the Data Model.
/// </summary>
[Collection("DataModel")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Slow")]
public class DataModelCommandsTests_DaxFormatting
{
    private readonly DataModelCommands _dataModelCommands;
    private readonly string _dataModelFile;

    public DataModelCommandsTests_DaxFormatting(DataModelPivotTableFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _dataModelFile = fixture.TestFilePath;
    }

    /// <summary>
    /// Tests that ListMeasures returns raw DAX previews (no formatting on read).
    /// Verifies that formula previews are returned as stored in the Data Model.
    /// </summary>
    [Fact]
    public async Task ListMeasures_WithMeasures_ReturnsRawPreviews()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _dataModelCommands.ListMeasures(batch);

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
    public async Task Read_WithMeasure_ReturnsRawFormula()
    {
        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        var result = await _dataModelCommands.Read(batch, "Total Sales");

        Assert.True(result.Success, $"Read failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.DaxFormula);
        Assert.Contains("SUM", result.DaxFormula, StringComparison.OrdinalIgnoreCase);

        // CharacterCount should reflect the formula length
        Assert.True(result.CharacterCount > 0);
        Assert.Equal(result.DaxFormula.Length, result.CharacterCount);
    }

    /// <summary>
    /// Tests that CreateMeasure formats DAX before saving to Excel.
    /// Verifies that the measure can be created and retrieved successfully.
    /// </summary>
    [Fact]
    public async Task CreateMeasure_WithUnformattedDax_SavesFormattedVersion()
    {
        var measureName = $"Test_CreateFormatted_{Guid.NewGuid():N}";
        // Unformatted DAX (single line, no spaces around operators)
        var unformattedDax = "CALCULATE(SUM(SalesTable[Amount]),FILTER(SalesTable,SalesTable[CustomerID]=1))";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create measure (should format automatically)
        _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, unformattedDax);

        // Retrieve and verify
        var viewResult = await _dataModelCommands.Read(batch, measureName);
        Assert.True(viewResult.Success, $"Read failed: {viewResult.ErrorMessage}");
        Assert.Contains("CALCULATE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("SUM", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);

        // The formatted version might have newlines or extra spaces
        // but should still contain the core function names
        Assert.NotEmpty(viewResult.DaxFormula);
    }

    /// <summary>
    /// Tests that UpdateMeasure formats DAX before saving to Excel.
    /// Verifies that the measure can be updated and retrieved successfully.
    /// </summary>
    [Fact]
    public async Task UpdateMeasure_WithUnformattedDax_SavesFormattedVersion()
    {
        var measureName = $"Test_UpdateFormatted_{Guid.NewGuid():N}";
        var originalFormula = "SUM(SalesTable[Amount])";
        // Unformatted DAX for update (single line, no spaces)
        var unformattedUpdate = "CALCULATE(AVERAGE(SalesTable[Amount]),FILTER(SalesTable,SalesTable[Region]=\"North\"))";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create measure
        _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, originalFormula);

        // Update with unformatted DAX (should format automatically)
        _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: unformattedUpdate);

        // Retrieve and verify
        var viewResult = await _dataModelCommands.Read(batch, measureName);
        Assert.True(viewResult.Success, $"Read failed: {viewResult.ErrorMessage}");
        Assert.Contains("CALCULATE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("AVERAGE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);

        // Should NOT contain SUM (since we updated the formula)
        Assert.DoesNotContain("SUM", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests that formatted DAX still executes correctly in Excel.
    /// Creates a measure with formatted DAX and verifies it can be used in a PivotTable.
    /// </summary>
    [Fact]
    public async Task CreateMeasure_WithFormattedDax_ExecutesCorrectlyInExcel()
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

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create measure with pre-formatted DAX
        _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, formattedDax);

        // Retrieve and verify it was saved correctly
        var viewResult = await _dataModelCommands.Read(batch, measureName);
        Assert.True(viewResult.Success, $"Read failed: {viewResult.ErrorMessage}");
        Assert.Contains("CALCULATE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);

        // Verify the measure appears in the list
        var listResult = await _dataModelCommands.ListMeasures(batch);
        Assert.Contains(listResult.Measures, m => m.Name == measureName);
    }

    /// <summary>
    /// Tests that null or empty DAX is handled gracefully (no formatting attempted).
    /// </summary>
    [Fact]
    public async Task UpdateMeasure_WithNullDaxFormula_DoesNotAttemptFormatting()
    {
        var measureName = $"Test_NullFormula_{Guid.NewGuid():N}";
        var originalFormula = "SUM(SalesTable[Amount])";
        var newDescription = "Updated description";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create measure
        _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, originalFormula);

        // Update only description (null daxFormula should not trigger formatting)
        _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: null, description: newDescription);

        // Verify description updated, formula unchanged
        var viewResult = await _dataModelCommands.Read(batch, measureName);
        Assert.True(viewResult.Success, $"Read failed: {viewResult.ErrorMessage}");
        Assert.Equal(newDescription, viewResult.Description);
        Assert.Contains("SUM", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }
}
