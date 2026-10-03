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

        RequireSuccess(result);
        Assert.Equal("SpacePivot", result.PivotTableName);
        Assert.Equal(4, result.AvailableFields.Count);
        AssertSpecialSheetPivot(result, "Sales Data", "A1:D6", "Sales Data", "F1", 335, 325);
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

        RequireSuccess(result);
        Assert.Equal("CrossSheetPivot", result.PivotTableName);
        AssertSpecialSheetPivot(result, "Sales Data", "A1:D6", "Pivot Output", "A1", 335, 325);
    }

    [Fact]
    public void CreateFromRange_SourceSheetWithHyphen_CreatesPivotTable()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Q1-Sales");
        RequireSuccess(_commands.SetValues(
            batch,
            "Q1-Sales",
            "A1:D3",
            [
                ["Region", "Product", "Sales", "Date"],
                ["North", "Widget", 100, "2025-01-15"],
                ["South", "Gadget", 200, "2025-02-10"],
            ]));
        RequireSuccess(_commands.SetNumberFormat(
            batch,
            "Q1-Sales",
            "D2:D3",
            "m/d/yyyy"));

        var result = _pivotCommands.CreateFromRange(
            batch,
            "Q1-Sales",
            "A1:D3",
            "Q1-Sales",
            "F1",
            "HyphenPivot");

        RequireSuccess(result);
        Assert.Equal("HyphenPivot", result.PivotTableName);
        AssertSpecialSheetPivot(result, "Q1-Sales", "A1:D3", "Q1-Sales", "F1", 100, 200);
    }

    private void CreateStandardSalesSheet(IExcelBatch batch, string sheetName)
    {
        _fixture.CreateNamedTestSheet(batch, sheetName);
        RequireSuccess(_commands.SetValues(
            batch,
            sheetName,
            "A1:D6",
            [
                ["Region", "Product", "Sales", "Date"],
                ["North", "Widget", 110, "2025-01-15"],
                ["North", "Widget", 150, "2025-01-20"],
                ["South", "Gadget", 200, "2025-02-10"],
                ["North", "Gadget", 75, "2025-02-15"],
                ["South", "Widget", 125, "2025-03-05"],
            ]));
        RequireSuccess(_commands.SetNumberFormat(
            batch,
            sheetName,
            "D2:D6",
            "m/d/yyyy"));
    }

    private void AssertSpecialSheetPivot(
        Sbroenne.ExcelMcp.Core.Models.PivotTableCreateResult result,
        string sourceSheet, string sourceRange, string destinationSheet, string position, int north, int south)
    {
        RequireSuccess(result);
        Assert.Equal(destinationSheet, result.SheetName);
        var before = RequireSuccess(_commands.GetValues(
            _fixture.BatchToken, sourceSheet, sourceRange));
        ReadNativePivot(destinationSheet, result.PivotTableName, pivot =>
        {
            Microsoft.Office.Interop.Excel.PivotCache? cache = null;
            Microsoft.Office.Interop.Excel.Range? range = null;
            Microsoft.Office.Interop.Excel.Range? cells = null;
            Microsoft.Office.Interop.Excel.Range? first = null;
            try
            {
                cache = pivot.PivotCache();
                Assert.Contains(sourceSheet, Convert.ToString(cache.SourceData),
                    StringComparison.Ordinal);
                Assert.Equal(before.Values.Count - 1, cache.RecordCount);
                range = pivot.TableRange2;
                Assert.Equal(result.Range, range.Address);
                cells = range.Cells;
                first = (Microsoft.Office.Interop.Excel.Range)cells[1, 1];
                Assert.Equal(position, first.Address[false, false]);
                return 0;
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref first);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref cells);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref range);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref cache);
            }
        });
        RequireSuccess(_pivotCommands.AddRowField(_fixture.BatchToken, result.PivotTableName, "Region"));
        RequireSuccess(_pivotCommands.AddValueField(_fixture.BatchToken, result.PivotTableName, "Sales"));
        RequireSuccess(_pivotCommands.Refresh(_fixture.BatchToken, result.PivotTableName));
        AssertPivotSales(north, south, result.PivotTableName);
        var after = RequireSuccess(_commands.GetValues(
            _fixture.BatchToken, sourceSheet, sourceRange));
        Assert.Equal(before.Values.Count, after.Values.Count);
        for (var index = 0; index < before.Values.Count; index++)
        {
            Assert.Equal(before.Values[index], after.Values[index]);
        }
        AssertOriginalSales();
    }
}
