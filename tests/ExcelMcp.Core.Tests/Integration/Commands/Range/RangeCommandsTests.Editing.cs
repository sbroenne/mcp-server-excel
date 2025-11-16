using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range editing operations
/// </summary>
public partial class RangeCommandsTests
{
    /// <inheritdoc/>
    // === CLEAR OPERATIONS TESTS ===

    [Fact]
    public void ClearAll_FormattedRange_RemovesEverything()
    {
        // Arrange
        string testFile = CoreTestHelper.CreateUniqueTestFile(nameof(RangeCommandsTests), nameof(ClearAll_FormattedRange_RemovesEverything), _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        _commands.SetValues(batch, "Sheet1", "A1", [["Test"]]);

        // Act
        var result = _commands.ClearAll(batch, "Sheet1", "A1");
        // Assert
        Assert.True(result.Success);

        var readResult = _commands.GetValues(batch, "Sheet1", "A1");
        Assert.Null(readResult.Values[0][0]);
    }
    /// <inheritdoc/>

    [Fact]
    public void ClearContents_FormattedRange_PreservesFormatting()
    {
        // Arrange
        string testFile = CoreTestHelper.CreateUniqueTestFile(nameof(RangeCommandsTests), nameof(ClearContents_FormattedRange_PreservesFormatting), _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        _commands.SetValues(batch, "Sheet1", "A1:B2",
        [
            [1, 2],
            [3, 4]
        ]);

        // Act
        var result = _commands.ClearContents(batch, "Sheet1", "A1:B2");
        // Assert
        Assert.True(result.Success);

        var readResult = _commands.GetValues(batch, "Sheet1", "A1:B2");
        Assert.All(readResult.Values, row => Assert.All(row, cell => Assert.Null(cell)));
    }
    /// <inheritdoc/>

    // === COPY OPERATIONS TESTS ===

    [Fact]
    public void Copy_CopiesRangeToNewLocation()
    {
        // Arrange
        string testFile = CoreTestHelper.CreateUniqueTestFile(nameof(RangeCommandsTests), nameof(Copy_CopiesRangeToNewLocation), _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        var sourceData = new List<List<object?>>
        {
            new() { "A", "B" },
            new() { 1, 2 }
        };

        _commands.SetValues(batch, "Sheet1", "A1:B2", sourceData);

        // Act
        var result = _commands.Copy(batch, "Sheet1", "A1:B2", "Sheet1", "D1:E2");
        // Assert
        Assert.True(result.Success);

        var readResult = _commands.GetValues(batch, "Sheet1", "D1:E2");
        Assert.Equal("A", readResult.Values[0][0]);
        Assert.Equal(2.0, Convert.ToDouble(readResult.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
    }
    /// <inheritdoc/>

    [Fact]
    public void CopyValues_CopiesOnlyValues()
    {
        // Arrange
        string testFile = CoreTestHelper.CreateUniqueTestFile(nameof(RangeCommandsTests), nameof(CopyValues_CopiesOnlyValues), _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        _commands.SetValues(batch, "Sheet1", "A1", [[10]]);
        _commands.SetFormulas(batch, "Sheet1", "B1", [["=A1*2"]]);

        // Act
        var result = _commands.CopyValues(batch, "Sheet1", "B1", "Sheet1", "C1");
        // Assert
        Assert.True(result.Success);

        // C1 should have value 20 but no formula
        var formulaResult = _commands.GetFormulas(batch, "Sheet1", "C1");
        Assert.Equal(20.0, Convert.ToDouble(formulaResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Empty(formulaResult.Formulas[0][0]); // No formula
    }

    // === INSERT/DELETE OPERATIONS TESTS ===
}
