using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    public void CreateFromRange_SourceSheetWithSpaces_CreatesPivotTable()
    {
        var batch = _fixture.BatchToken;
        CreateStandardSalesSheet(batch, "Sales Data");

        var result = _pivotCommands.CreateFromRange(
            batch,
            "Sales Data",
            "A1:D6",
            "Sales Data",
            "F1",
            "SpacePivot");

        Assert.True(
            result.Success,
            $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SpacePivot", result.PivotTableName);
        Assert.Equal(4, result.AvailableFields.Count);
    }

    [Fact]
    public void CreateFromRange_DestinationSheetWithSpaces_CreatesPivotTable()
    {
        var batch = _fixture.BatchToken;
        CreateStandardSalesSheet(batch, "Sales Data");
        _fixture.CreateNamedTestSheet(batch, "Pivot Output");

        var result = _pivotCommands.CreateFromRange(
            batch,
            "Sales Data",
            "A1:D6",
            "Pivot Output",
            "A1",
            "CrossSheetPivot");

        Assert.True(
            result.Success,
            $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("CrossSheetPivot", result.PivotTableName);
    }

    [Fact]
    public void CreateFromRange_SourceSheetWithHyphen_CreatesPivotTable()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Q1-Sales");
        _commands.SetValues(
            batch,
            "Q1-Sales",
            "A1:D3",
            [
                ["Region", "Product", "Sales", "Date"],
                ["North", "Widget", 100, "2025-01-15"],
                ["South", "Gadget", 200, "2025-02-10"],
            ]);
        _commands.SetNumberFormat(
            batch,
            "Q1-Sales",
            "D2:D3",
            "m/d/yyyy");

        var result = _pivotCommands.CreateFromRange(
            batch,
            "Q1-Sales",
            "A1:D3",
            "Q1-Sales",
            "F1",
            "HyphenPivot");

        Assert.True(
            result.Success,
            $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("HyphenPivot", result.PivotTableName);
    }

    private void CreateStandardSalesSheet(IExcelBatch batch, string sheetName)
    {
        _fixture.CreateNamedTestSheet(batch, sheetName);
        _commands.SetValues(
            batch,
            sheetName,
            "A1:D6",
            [
                ["Region", "Product", "Sales", "Date"],
                ["North", "Widget", 100, "2025-01-15"],
                ["North", "Widget", 150, "2025-01-20"],
                ["South", "Gadget", 200, "2025-02-10"],
                ["North", "Gadget", 75, "2025-02-15"],
                ["South", "Widget", 125, "2025-03-05"],
            ]);
        _commands.SetNumberFormat(
            batch,
            sheetName,
            "D2:D6",
            "m/d/yyyy");
    }
}
