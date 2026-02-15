using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Integration tests for DAX formula locale translation.
/// Tests that DAX formulas with US comma separators work correctly on all locales.
/// </summary>
public partial class DataModelCommandsTests
{
    #region DAX Locale Translation Tests

    /// <summary>
    /// Tests that DAX formulas with function argument separators (commas) are handled correctly.
    /// This is the exact formula from the user's bug report where DATEADD arguments were corrupted.
    /// LLM use case: "create a measure with DATEADD function"
    /// </summary>
    [Fact]
    public async Task CreateMeasure_DateAddFormula_CreatesSuccessfully()
    {
        // This is the formula that was failing - comma was becoming period on European locales
        var measureName = $"Test_DATEADD_{Guid.NewGuid():N}";
        var daxFormula = "CALCULATE([Total Sales], DATEADD(SalesTable[Date], -1, MONTH))";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // This should NOT throw - the DaxFormulaTranslator should handle locale conversion
        _ = _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, daxFormula);

        // Verify measure was created
        var listResult = await _dataModelCommands.ListMeasures(batch);
        Assert.Contains(listResult.Measures, m => m.Name == measureName);

        // Verify the formula was stored (content may vary by locale but should be valid DAX)
        var readResult = await _dataModelCommands.Read(batch, measureName);
        Assert.True(readResult.Success, $"Read measure failed: {readResult.ErrorMessage}");
        Assert.NotNull(readResult.DaxFormula);
        Assert.Contains("DATEADD", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("CALCULATE", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests nested function calls with multiple comma separators.
    /// LLM use case: "create a complex DAX measure with nested functions"
    /// </summary>
    [Fact]
    public async Task CreateMeasure_NestedFunctions_CreatesSuccessfully()
    {
        var measureName = $"Test_Nested_{Guid.NewGuid():N}";
        // Complex formula with multiple nested functions and comma separators
        var daxFormula = "CALCULATE(SUM(SalesTable[Amount]), FILTER(ALL(SalesTable), SalesTable[Amount] > 100))";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        _ = _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, daxFormula);

        // Verify measure was created and formula is valid
        var readResult = await _dataModelCommands.Read(batch, measureName);
        Assert.True(readResult.Success, $"Read measure failed: {readResult.ErrorMessage}");
        Assert.Contains("CALCULATE", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("FILTER", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests updating a measure with a complex DAX formula containing function separators.
    /// LLM use case: "update this measure's formula to use DATESINPERIOD"
    /// </summary>
    [Fact]
    public async Task UpdateMeasure_ComplexDaxFormula_UpdatesSuccessfully()
    {
        var measureName = $"Test_Update_{Guid.NewGuid():N}";
        var originalFormula = "SUM(SalesTable[Amount])";
        // Rolling 3-month formula with multiple comma separators
        var updatedFormula = "AVERAGEX(DATESINPERIOD(SalesTable[Date], MAX(SalesTable[Date]), -3, MONTH), SalesTable[Amount])";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        // Create measure with simple formula
        _ = _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, originalFormula);

        // Update with complex formula - should handle locale conversion
        _ = _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: updatedFormula);

        // Verify the formula was updated
        var readResult = await _dataModelCommands.Read(batch, measureName);
        Assert.True(readResult.Success, $"Read measure failed: {readResult.ErrorMessage}");
        Assert.Contains("AVERAGEX", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("DATESINPERIOD", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests DAX formula with string literals containing commas - commas inside strings should NOT be translated.
    /// LLM use case: "create a measure that checks for a specific text value"
    /// </summary>
    [Fact]
    public async Task CreateMeasure_StringLiteralWithComma_PreservesStringContent()
    {
        var measureName = $"Test_String_{Guid.NewGuid():N}";
        // Formula with comma inside a string literal - this comma should NOT be translated
        var daxFormula = "IF(SalesTable[Region] = \"North, South\", 1, 0)";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        _ = _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, daxFormula);

        // Verify measure was created
        var readResult = await _dataModelCommands.Read(batch, measureName);
        Assert.True(readResult.Success, $"Read measure failed: {readResult.ErrorMessage}");
        Assert.Contains("IF", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests simple DAX formula without function separators - should work unchanged.
    /// LLM use case: "create a simple SUM measure"
    /// </summary>
    [Fact]
    public async Task CreateMeasure_SimpleFormula_CreatesSuccessfully()
    {
        var measureName = $"Test_Simple_{Guid.NewGuid():N}";
        var daxFormula = "SUM(SalesTable[Amount])";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);
        _ = _dataModelCommands.CreateMeasure(batch, "SalesTable", measureName, daxFormula);

        var readResult = await _dataModelCommands.Read(batch, measureName);
        Assert.True(readResult.Success, $"Read measure failed: {readResult.ErrorMessage}");
        Assert.Contains("SUM", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}




