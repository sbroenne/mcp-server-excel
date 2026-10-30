using System.Diagnostics;
using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Slow")]
public sealed class PersistentServicePivotTableCacheScaleTests(
    PersistentServiceWorkbookFixture fixture,
    ITestOutputHelper output) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void ShareExistingCache_Handles200000DistinctItemsBeforeSessionTimeout()
    {
        const int initialRecords = 200_000;
        var batch = _fixture.BatchToken;
        var source = _fixture.CreateTestSheet(batch);
        var target = _fixture.CreateTestSheet(batch);
        var moved = _fixture.CreateTestSheet(batch);
        var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var tables = _fixture.CreateCommands<IPersistentTableCommands>();

        var values = new object[initialRecords + 1, 4];
        values[0, 0] = "SaleId";
        values[0, 1] = "Region";
        values[0, 2] = "Quantity";
        values[0, 3] = "NetEUR";
        string[] regions = ["North", "South", "East", "West"];
        for (int row = 1; row <= initialRecords; row++)
        {
            values[row, 0] = row;
            values[row, 1] = regions[(row - 1) % regions.Length];
            values[row, 2] = 1;
            values[row, 3] = 10;
        }
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, source);
                Assert.NotNull(sheet);
                range = sheet.Range[$"A1:D{initialRecords + 1}"];
                range.Value2 = values;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        RequireSuccess(tables.Create(batch, source, "LargeSales", $"A1:D{initialRecords + 1}", true));
        RequireSuccess(pivots.CreateFromTable(batch, "LargeSales", target, "A1", "RegionRevenue"));
        RequireSuccess(pivots.CreateFromTable(batch, "LargeSales", moved, "A1", "RegionUnits"));
        RequireSuccess(pivots.AddRowField(batch, "RegionRevenue", "Region"));
        RequireSuccess(pivots.AddValueField(batch, "RegionRevenue", "NetEUR"));
        RequireSuccess(pivots.AddRowField(batch, "RegionUnits", "Region"));
        RequireSuccess(pivots.AddValueField(batch, "RegionUnits", "Quantity"));
        RequireSuccess(_commands.SetValues(batch, source, $"A{initialRecords + 2}:D{initialRecords + 4}",
            [[initialRecords + 1, "North", 1, 10],
             [initialRecords + 2, "South", 1, 10],
             [initialRecords + 3, "East", 1, 10]], overwritePolicy: OverwritePolicy.Allow));
        RequireSuccess(tables.Resize(batch, "LargeSales", $"A1:D{initialRecords + 4}"));
        RequireSuccess(pivots.Refresh(batch, "RegionRevenue"));

        var targetBefore = ReadCacheState(target, "RegionRevenue");
        var movedBefore = ReadCacheState(moved, "RegionUnits");
        Assert.NotEqual(targetBefore.Index, movedBefore.Index);
        Assert.Equal((initialRecords + 3, initialRecords + 3), (targetBefore.Records, targetBefore.Members));
        Assert.Equal((initialRecords, initialRecords), (movedBefore.Records, movedBefore.Members));
        AssertUnits(pivots, batch, false);

        var stopwatch = Stopwatch.StartNew();
        var result = RequireSuccess(pivots.ShareCache(batch, moved, "RegionUnits", "RegionRevenue", target));
        stopwatch.Stop();
        output.WriteLine($"Sharing completed in {stopwatch.Elapsed.TotalSeconds:F2} seconds.");

        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(120),
            $"Sharing 200,000 cached items took {stopwatch.Elapsed}, exceeding the 120-second session timeout.");
        var targetAfter = ReadCacheState(target, "RegionRevenue");
        var movedAfter = ReadCacheState(moved, "RegionUnits");
        Assert.True(result.Changed);
        Assert.Equal(movedBefore.Index, result.PreviousCacheIndex);
        Assert.Equal(targetAfter.Index, result.CacheIndex);
        Assert.Equal(targetAfter.Index, result.TargetCacheIndex);
        Assert.Equal(targetAfter.Index, movedAfter.Index);
        Assert.Equal((initialRecords + 3, initialRecords + 3), (movedAfter.Records, movedAfter.Members));
        Assert.Empty(result.PreviousCacheUsers);
        Assert.Equal(2, result.SharedPivotTables.Count);
        Assert.Contains(result.SharedPivotTables, item => item.SheetName == target && item.PivotTableName == "RegionRevenue");
        Assert.Contains(result.SharedPivotTables, item => item.SheetName == moved && item.PivotTableName == "RegionUnits");
        AssertUnits(pivots, batch, true);
    }

    private static void AssertUnits(IPersistentPivotTableCommands pivots, IExcelBatch batch,
        bool includesNewRows)
    {
        var values = RequireSuccess(pivots.GetData(batch, "RegionUnits")).Values;
        var expected = new Dictionary<string, int>
        {
            ["North"] = 50_000 + (includesNewRows ? 1 : 0),
            ["South"] = 50_000 + (includesNewRows ? 1 : 0),
            ["East"] = 50_000 + (includesNewRows ? 1 : 0),
            ["West"] = 50_000
        };
        Assert.Equal(expected.Count + 2, values.Count);
        foreach (var (region, count) in expected)
        {
            var row = Assert.Single(values, item => string.Equals(item[0]?.ToString(), region, StringComparison.Ordinal));
            Assert.Equal(count, Convert.ToInt32(row[1], CultureInfo.InvariantCulture));
        }
        Assert.Equal(200_000 + (includesNewRows ? 3 : 0),
            Convert.ToInt32(values[^1][1], CultureInfo.InvariantCulture));
    }

    private (int Index, int Records, int Members) ReadCacheState(string sheetName, string pivotName) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? cache = null;
            Excel.PivotField? field = null;
            Excel.PivotItems? items = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                pivots = sheet.PivotTables();
                pivot = pivots.Item(pivotName);
                cache = pivot.PivotCache();
                field = (Excel.PivotField)pivot.PivotFields("SaleId");
                items = (Excel.PivotItems)field.PivotItems();
                return (pivot.CacheIndex, cache.RecordCount, items.Count);
            }
            finally
            {
                ComUtilities.Release(ref items);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
            }
        });
}
