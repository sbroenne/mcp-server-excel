using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Integration tests for Range daemon handlers.
/// Verifies that daemon handlers correctly delegate to Core Commands.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
public class RangeDaemonHandlerTests : DaemonIntegrationTestBase
{
    private readonly RangeCommands _rangeCommands = new();
    private readonly SheetCommands _sheetCommands = new();

    public RangeDaemonHandlerTests(TempDirectoryFixture fixture) : base(fixture) { }

    [Fact]
    [Trait("Speed", "Fast")]
    public void RangeGetValues_EmptyRange_ReturnsEmptyValues()
    {
        // Arrange
        using var batch = CreateBatch();
        var sheets = _sheetCommands.List(batch);
        var sheetName = sheets.Worksheets.First().Name;

        // Act
        var result = _rangeCommands.GetValues(batch, sheetName, "A1:B2");

        // Assert
        Assert.True(result.Success, $"GetValues failed: {result.ErrorMessage}");
        Assert.NotNull(result.Values);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void RangeSetValues_ValidData_WritesValues()
    {
        // Arrange
        using var batch = CreateBatch();
        var sheets = _sheetCommands.List(batch);
        var sheetName = sheets.Worksheets.First().Name;
        var values = new List<List<object?>>
        {
            new() { "Header1", "Header2" },
            new() { "Value1", 123 }
        };

        // Act
        var setResult = _rangeCommands.SetValues(batch, sheetName, "A1:B2", values);

        // Assert
        Assert.True(setResult.Success, $"SetValues failed: {setResult.ErrorMessage}");

        // Verify by reading back
        var getResult = _rangeCommands.GetValues(batch, sheetName, "A1:B2");
        Assert.True(getResult.Success);
        Assert.NotNull(getResult.Values);
        Assert.Equal(2, getResult.Values.Count);
        Assert.Equal("Header1", getResult.Values[0][0]?.ToString());
        Assert.Equal("Header2", getResult.Values[0][1]?.ToString());
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void RangeGetUsedRange_AfterWritingData_ReturnsCorrectRange()
    {
        // Arrange
        using var batch = CreateBatch();
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var sheetName = $"Range{uniqueId}";
        _sheetCommands.Create(batch, sheetName);

        // Write some data
        var values = new List<List<object?>>
        {
            new() { "A", "B", "C" },
            new() { 1, 2, 3 }
        };
        _rangeCommands.SetValues(batch, sheetName, "A1:C2", values);

        // Act
        var result = _rangeCommands.GetUsedRange(batch, sheetName);

        // Assert
        Assert.True(result.Success, $"GetUsedRange failed: {result.ErrorMessage}");
        Assert.Contains("$A$1", result.RangeAddress); // Absolute address format
    }
}
