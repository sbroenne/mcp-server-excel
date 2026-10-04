using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for range editing operations
/// </summary>
public sealed partial class PersistentServiceRangeEditingTests
{
    // === CLEAR OPERATIONS TESTS ===

    [Fact]
    public void ClearAll_FormattedRange_RemovesEverything()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [["Test"]]));
        var normal = PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "B1");
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "A1", "0.00%").Success);
        RequireSuccess(_commands.Format(batch, sheetName, ["A1"],
            new() { Bold = true, FillColor = "#FF0000" }));
        Assert.NotEqual(normal, PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "A1"));

        // Act
        var result = _commands.ClearAll(batch, sheetName, "A1");
        // Assert
        Assert.True(result.Success);

        var readResult = _commands.GetValues(batch, sheetName, "A1");
        Assert.Null(readResult.Values[0][0]);
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.Equal(normal, PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "A1"));
    }

    [Fact]
    public void ClearContents_FormattedRange_PreservesFormatting()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:B2",
        [
            [1, 2],
            [3, 4]
        ]));
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "A1:B2", "0.00%").Success);
        RequireSuccess(_commands.Format(batch, sheetName, ["A1:B2"],
            new() { Bold = true, FillColor = "#FF0000" }));
        var before = PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, "A1");

        // Act
        var result = _commands.ClearContents(batch, sheetName, "A1:B2");
        // Assert
        Assert.True(result.Success);

        var readResult = _commands.GetValues(batch, sheetName, "A1:B2");
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.All(readResult.Values, row => Assert.All(row, cell => Assert.Null(cell)));
        foreach (var cell in new[] { "A1", "B1", "A2", "B2" })
            Assert.Equal(before, PersistentServiceRangeVerification.ReadFormat(_fixture, sheetName, cell));
    }

    // === COPY OPERATIONS TESTS ===

    [Fact]
    public void Copy_CopiesRangeToNewLocation()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var sourceData = new List<List<object?>>
        {
            new() { "A", "B" },
            new() { 1, 2 }
        };

        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:B2", sourceData));

        // Act
        var result = _commands.Copy(batch, sheetName, "A1:B2", sheetName, "D1:E2", PasteKind.All);
        // Assert
        Assert.True(result.Success);

        var readResult = _commands.GetValues(batch, sheetName, "D1:E2");
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.Equal("A", readResult.Values[0][0]);
        Assert.Equal("B", readResult.Values[0][1]);
        Assert.Equal(1d, Convert.ToDouble(readResult.Values[1][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(2.0, Convert.ToDouble(readResult.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
        var unchanged = _commands.GetValues(batch, sheetName, "A1:B2");
        Assert.True(unchanged.Success, unchanged.ErrorMessage);
        Assert.Equal(readResult.Values.SelectMany(row => row), unchanged.Values.SelectMany(row => row));
    }

    [Fact]
    public void CopyValues_CopiesOnlyValues()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [[10]]));
        RequireSuccess(_commands.SetFormulas(batch, sheetName, "B1", [["=A1*2"]]));

        // Act
        var result = _commands.Copy(batch, sheetName, "B1", sheetName, "C1", PasteKind.Values);
        // Assert
        Assert.True(result.Success);

        // C1 should have value 20 but no formula
        var formulaResult = RequireSuccess(_commands.GetFormulas(batch, sheetName, "C1"));
        Assert.Equal(20.0, Convert.ToDouble(formulaResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Empty(formulaResult.Formulas[0][0]); // No formula
        var source = RequireSuccess(_commands.GetFormulas(batch, sheetName, "A1:B1"));
        Assert.Equal("=A1*2", source.Formulas[0][1]);
        Assert.Equal(10, source.Values[0][0]);
        Assert.Equal(20, source.Values[0][1]);
    }

    // === INSERT/DELETE OPERATIONS TESTS ===
}
