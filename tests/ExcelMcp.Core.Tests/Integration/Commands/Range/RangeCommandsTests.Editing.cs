using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range editing operations
/// </summary>
public partial class RangeCommandsTests
{
    // === CLEAR OPERATIONS TESTS ===

    [Fact]
    public void ClearAll_FormattedRange_RemovesEverything()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1", [["Test"]]);

        // Act
        var result = _commands.ClearAll(batch, sheetName, "A1");
        // Assert
        Assert.True(result.Success);

        var readResult = _commands.GetValues(batch, sheetName, "A1");
        Assert.Null(readResult.Values[0][0]);
    }

    [Fact]
    public void ClearContents_FormattedRange_PreservesFormatting()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1:B2",
        [
            [1, 2],
            [3, 4]
        ]);

        // Act
        var result = _commands.ClearContents(batch, sheetName, "A1:B2");
        // Assert
        Assert.True(result.Success);

        var readResult = _commands.GetValues(batch, sheetName, "A1:B2");
        Assert.All(readResult.Values, row => Assert.All(row, cell => Assert.Null(cell)));
    }

    // === COPY OPERATIONS TESTS ===

    [Fact]
    public void Copy_CopiesRangeToNewLocation()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var sourceData = new List<List<object?>>
        {
            new() { "A", "B" },
            new() { 1, 2 }
        };

        _commands.SetValues(batch, sheetName, "A1:B2", sourceData);

        // Act
        var result = _commands.Copy(batch, sheetName, "A1:B2", sheetName, "D1:E2");
        // Assert
        Assert.True(result.Success);

        var readResult = _commands.GetValues(batch, sheetName, "D1:E2");
        Assert.Equal("A", readResult.Values[0][0]);
        Assert.Equal(2.0, Convert.ToDouble(readResult.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void CopyValues_CopiesOnlyValues()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1", [[10]]);
        _commands.SetFormulas(batch, sheetName, "B1", [["=A1*2"]]);

        // Act
        var result = _commands.CopyValues(batch, sheetName, "B1", sheetName, "C1");
        // Assert
        Assert.True(result.Success);

        // C1 should have value 20 but no formula
        var formulaResult = _commands.GetFormulas(batch, sheetName, "C1");
        Assert.Equal(20.0, Convert.ToDouble(formulaResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Empty(formulaResult.Formulas[0][0]); // No formula
    }

    // === INSERT/DELETE OPERATIONS TESTS ===
}




