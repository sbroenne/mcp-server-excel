using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Theory]
    [InlineData("H3", true)]
    [InlineData("H30", false)]
    [Trait("Speed", "Medium")]
    [Trait("Category", "Regular")]
    public async Task AddValueField_WithAdjacentPivot_ReportsActualLayout(string otherLocation, bool blocked)
    {
        var batch = _fixture.BatchToken;
        var sourceSheet = _fixture.CreateTestSheet(batch);
        var destinationSheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sourceSheet, "A1:I9",
        [
            ["Group", "Status", "V1", "V2", "V3", "V4", "V5", "V6", "V7"],
            ["A", "Open", 1, 2, 3, 4, 5, 6, 7],
            ["A", "Closed", 1, 2, 3, 4, 5, 6, 7],
            ["A", "Open", 1, 2, 3, 4, 5, 6, 7],
            ["B", "Open", 1, 2, 3, 4, 5, 6, 7],
            ["B", "Closed", 1, 2, 3, 4, 5, 6, 7],
            ["B", "Closed", 1, 2, 3, 4, 5, 6, 7],
            ["C", "Open", 1, 2, 3, 4, 5, 6, 7],
            ["C", "Closed", 1, 2, 3, 4, 5, 6, 7],
        ]));
        RequireSuccess(_tableCommands.Create(batch, sourceSheet, "tblData", "A1:I9", true));
        var sourceBefore = RequireSuccess(_commands.GetValues(batch, sourceSheet, "A1:I9")).Values;
        RequireSuccess(_pivotCommands.CreateFromTable(batch, "tblData", destinationSheet, "A3", "ptC"));
        RequireSuccess(_pivotCommands.CreateFromTable(batch, "tblData", destinationSheet, otherLocation, "ptD"));
        var otherBefore = RequireSuccess(_pivotCommands.Read(batch, "ptD")).PivotTable;
        Assert.Equal(0, otherBefore.ValueFieldCount);
        Assert.Null(otherBefore.LayoutError);

        RequireSuccess(_pivotCommands.AddColumnField(batch, "ptC", "Status"));
        for (var index = 1; index <= 2; index++)
        {
            RequireSuccess(_pivotCommands.AddValueField(batch, "ptC", $"V{index}",
                Sbroenne.ExcelMcp.Core.Models.AggregationFunction.Sum, $"Total V{index}"));
            var info = RequireSuccess(_pivotCommands.Read(batch, "ptC")).PivotTable;
            Assert.Equal(index, info.ValueFieldCount);
            Assert.Contains(":", info.Range, StringComparison.Ordinal);
            Assert.Null(info.LayoutError);
        }

        if (blocked)
        {
            var result = await _fixture.SendForFailureAsync("pivottablefield.add-value-field",
                new { pivotTableName = "ptC", fieldName = "V3", aggregationFunction = "Sum", customName = "Total V3" });
            Assert.False(result.Success);
            Assert.Contains("#SPILL!", result.ErrorMessage, StringComparison.Ordinal);
            Assert.Contains("not been rolled back", result.ErrorMessage, StringComparison.Ordinal);
        }
        else
        {
            var result = _pivotCommands.AddValueField(batch, "ptC", "V3",
                Sbroenne.ExcelMcp.Core.Models.AggregationFunction.Sum, "Total V3");
            RequireSuccess(result);
            Assert.Equal("Total V3", result.CustomName);
        }

        var after = RequireSuccess(_pivotCommands.Read(batch, "ptC")).PivotTable;
        Assert.Equal(3, after.ValueFieldCount);
        if (blocked)
        {
            Assert.Equal("$A$3", after.Range);
            Assert.Contains("#SPILL!", after.LayoutError, StringComparison.Ordinal);
            ReadNativePivot(destinationSheet, "ptC", pivot =>
            {
                Microsoft.Office.Interop.Excel.Range? range = null;
                try
                {
                    range = pivot.TableRange1;
                    Assert.Equal("$A$3", range.Address);
                    Assert.Equal(-2146826243, Convert.ToInt32((object)range.Value2, System.Globalization.CultureInfo.InvariantCulture));
                    return 0;
                }
                finally
                {
                    Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref range);
                }
            });
        }
        else
        {
            Assert.Contains(":", after.Range, StringComparison.Ordinal);
            Assert.Null(after.LayoutError);
            var data = RequireSuccess(_pivotCommands.GetData(batch, "ptC"));
            Assert.Contains(data.Values.SelectMany(row => row),
                value => Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) == "24");
        }

        var listed = RequireSuccess(_pivotCommands.List(batch));
        var listedPivot = Assert.Single(listed.PivotTables, pivot => pivot.Name == "ptC");
        Assert.Equal(after.Range, listedPivot.Range);
        Assert.Equal(after.LayoutError, listedPivot.LayoutError);
        Assert.Equal(after.ValueFieldCount, listedPivot.ValueFieldCount);
        var otherAfter = RequireSuccess(_pivotCommands.Read(batch, "ptD")).PivotTable;
        Assert.Equal(otherBefore.Range, otherAfter.Range);
        Assert.Equal(0, otherAfter.ValueFieldCount);
        Assert.Null(otherAfter.LayoutError);
        var sourceAfter = RequireSuccess(_commands.GetValues(batch, sourceSheet, "A1:I9")).Values;
        Assert.Equal(sourceBefore.Count, sourceAfter.Count);
        for (var index = 0; index < sourceBefore.Count; index++)
        {
            Assert.Equal(sourceBefore[index], sourceAfter[index]);
        }
        AssertOriginalSales();
    }

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
                Assert.Contains(sourceSheet, Convert.ToString((object)cache.SourceData, System.Globalization.CultureInfo.InvariantCulture),
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
