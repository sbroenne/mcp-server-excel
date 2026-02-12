using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Calculation;

/// <summary>
/// Integration tests for calculation mode control (get-mode, set-mode, calculate actions).
/// Tests validate explicit control over Excel's automatic/manual calculation mode.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Feature", "CalculationMode")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public class CalculationModeCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly CalculationModeCommands _commands;
    private readonly TempDirectoryFixture _fixture;

    public CalculationModeCommandsTests(TempDirectoryFixture fixture)
    {
        _commands = new CalculationModeCommands();
        _fixture = fixture;
    }

    /// <summary>
    /// Verify get-mode returns current calculation state as automatic (default).
    /// </summary>
    [Fact]
    public void GetMode_ReturnsAutomaticByDefault()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.GetMode(batch);

        // Assert
        Assert.True(result.Success);
        Assert.Equal("automatic", result.Mode);
        Assert.Equal(-4105, result.ModeValue); // xlCalculationAutomatic
    }

    /// <summary>
    /// Verify set-mode can switch to manual.
    /// </summary>
    [Fact]
    public void SetMode_ToManual_Succeeds()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var setResult = _commands.SetMode(batch, CalculationMode.Manual);

        // Assert
        Assert.True(setResult.Success);

        // Verify it's actually manual
        var getResult = _commands.GetMode(batch);
        Assert.Equal("manual", getResult.Mode);
        Assert.Equal(-4135, getResult.ModeValue); // xlCalculationManual
    }

    /// <summary>
    /// Verify set-mode can switch to semi-automatic.
    /// </summary>
    [Fact]
    public void SetMode_ToSemiAutomatic_Succeeds()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var setResult = _commands.SetMode(batch, CalculationMode.SemiAutomatic);

        // Assert
        Assert.True(setResult.Success);

        // Verify it's actually semi-automatic
        var getResult = _commands.GetMode(batch);
        Assert.Equal("semi-automatic", getResult.Mode);
        Assert.Equal(2, getResult.ModeValue); // xlCalculationSemiautomatic
    }

    /// <summary>
    /// Verify calculate-workbook scope succeeds.
    /// </summary>
    [Fact]
    public void Calculate_WorkbookScope_Succeeds()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);

        // Switch to manual mode first
        _commands.SetMode(batch, CalculationMode.Manual);

        // Calculate workbook
        var result = _commands.Calculate(batch, CalculationScope.Workbook);

        // Assert
        Assert.True(result.Success);
    }

    /// <summary>
    /// Verify calculate-sheet scope requires sheet name.
    /// </summary>
    [Fact]
    public void Calculate_SheetScope_RequiresSheetName()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Calculate(batch, CalculationScope.Sheet, null);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("sheetName is required", result.ErrorMessage ?? "");
    }

    /// <summary>
    /// Verify calculate-sheet scope succeeds with sheet name.
    /// </summary>
    [Fact]
    public void Calculate_SheetScope_WithValidSheetName_Succeeds()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);

        // Switch to manual mode
        _commands.SetMode(batch, CalculationMode.Manual);

        // Calculate specific sheet
        var result = _commands.Calculate(batch, CalculationScope.Sheet, "Sheet1");

        // Assert
        Assert.True(result.Success);
    }

    /// <summary>
    /// Verify calculate-range scope requires both sheet and range.
    /// </summary>
    [Fact]
    public void Calculate_RangeScope_RequiresBothSheetAndRange()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.Calculate(batch, CalculationScope.Range, "Sheet1", null);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("rangeAddress are required", result.ErrorMessage ?? "");
    }

    /// <summary>
    /// Verify calculate-range scope succeeds with both parameters.
    /// </summary>
    [Fact]
    public void Calculate_RangeScope_WithValidParameters_Succeeds()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);

        // Switch to manual mode
        _commands.SetMode(batch, CalculationMode.Manual);

        // Calculate specific range
        var result = _commands.Calculate(batch, CalculationScope.Range, "Sheet1", "A1:C10");

        // Assert
        Assert.True(result.Success);
    }
}




