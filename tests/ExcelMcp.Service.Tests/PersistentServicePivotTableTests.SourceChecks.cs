using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Source checks that must happen before Excel creates a PivotTable.
/// </summary>
public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    public void CreateFromRange_SourceSheetNameWithApostrophe_CreatesPivot()
    {
        var batch = _fixture.BatchToken;
        var sourceSheet = $"Bob's {Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sourceSheet);
        RequireSuccess(_commands.SetValues(batch, sourceSheet, "A1:D6",
        [
            ["Region", "Product", "Sales", "Date"],
            ["North", "Widget", 100, "2025-01-15"],
            ["North", "Widget", 150, "2025-01-20"],
            ["South", "Gadget", 200, "2025-02-10"],
            ["North", "Gadget", 75, "2025-02-15"],
            ["South", "Widget", 125, "2025-03-05"],
        ]));

        var result = _pivotCommands.CreateFromRange(
            batch, sourceSheet, "A1:D6", _salesSheetName, "F1", "ApostrophePivot");

        RequireSuccess(result);
        Assert.Equal($"'{sourceSheet.Replace("'", "''", StringComparison.Ordinal)}'!A1:D6", result.SourceData);
        AssertCreatedPivot(result, _salesSheetName, "F1");
    }

    [Fact]
    public void CreateFromRange_BlankHeaderRow_RejectsBeforeCreatingPivot()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.SetValues(batch, _salesSheetName, "A9:B10",
        [
            ["North", 10],
            ["South", 20],
        ]));

        var error = Assert.ThrowsAny<Exception>(() =>
            _pivotCommands.CreateFromRange(batch, _salesSheetName, "A8:B10", _salesSheetName, "H1", "BlankHeaderPivot"));

        Assert.Contains("No field headers found", error.Message, StringComparison.Ordinal);
        Assert.Empty(RequireSuccess(_pivotCommands.List(batch)).PivotTables);
        AssertOriginalSales();
    }

    [Fact]
    public void CreateFromRange_PartiallyBlankHeaderRow_RejectsBeforeCreatingPivot()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_commands.SetValues(batch, _salesSheetName, "A9:C11",
        [
            ["Region", "", "Sales"],
            ["North", "Widget", 10],
            ["South", "Gadget", 20],
        ]));
        var namesBefore = RequireSuccess(_pivotCommands.List(batch)).PivotTables
            .Select(table => table.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var error = Assert.ThrowsAny<Exception>(() =>
            _pivotCommands.CreateFromRange(
                batch, _salesSheetName, "A9:C11", _salesSheetName, "H1", "PartialBlankHeaderPivot"));

        Assert.Contains("header", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(namesBefore, RequireSuccess(_pivotCommands.List(batch)).PivotTables
            .Select(table => table.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase));
        AssertOriginalSales();
    }
}
