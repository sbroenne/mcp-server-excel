using System.Dynamic;
using System.Reflection;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PivotTable;

// Service cannot inject a cancelled token into an already-running STA callback.
[Collection("Sequential")]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Slicer")]
[Trait("Speed", "Medium")]
public sealed class SlicerCancellationTests(TempDirectoryFixture fixture) :
    IClassFixture<TempDirectoryFixture>
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CreateSlicer_CancelledCallback_StopsBeforeCreation(bool existingCache)
    {
        using var innerBatch = ExcelSession.BeginBatch(fixture.CreateTestFile());
        var commands = new PivotTableCommands();
        PreparePivot(innerBatch, commands);
        if (existingCache)
        {
            var existing = commands.CreateSlicer(innerBatch, "CancellationPivot", "Region",
                "ExistingSlicer", "SlicerData", "E1");
            Assert.True(existing.Success, existing.ErrorMessage);
            var filtered = commands.SetSlicerSelection(innerBatch, "ExistingSlicer", ["North"]);
            Assert.True(filtered.Success, filtered.ErrorMessage);
            Assert.Equal(["North"], filtered.SelectedItems);
        }
        var initialCounts = GetSlicerCounts(innerBatch);
        Assert.Equal(existingCache ? (1, 1) : (0, 0), initialCounts);
        AssertPivot(innerBatch, existingCache ? 10 : 30,
            existingCache ? ["North"] : ["North", "South"]);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var batch = new InjectedCancellationBatch(innerBatch, cancellation.Token);

        Assert.ThrowsAny<OperationCanceledException>(() =>
            commands.CreateSlicer(batch, "CancellationPivot", "Region",
                "CancelledSlicer", "SlicerData", "E1"));
        Assert.Equal(initialCounts, GetSlicerCounts(innerBatch));
        AssertPivot(innerBatch, existingCache ? 10 : 30,
            existingCache ? ["North"] : ["North", "South"]);

        var followUp = commands.CreateSlicer(innerBatch, "CancellationPivot", "Region",
            "FollowUpSlicer", "SlicerData", "E1");
        Assert.True(followUp.Success, followUp.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(followUp.ErrorMessage));
        Assert.Equal(["North", "South"], followUp.AvailableItems.Order());
        Assert.Equal(existingCache ? ["North"] : ["North", "South"], followUp.SelectedItems.Order());
        Assert.Equal(existingCache ? (1, 2) : (1, 1), GetSlicerCounts(innerBatch));
        var replaced = commands.SetSlicerSelection(innerBatch, "FollowUpSlicer", ["South"]);
        Assert.True(replaced.Success, replaced.ErrorMessage);
        Assert.Equal(["South"], replaced.SelectedItems);
        AssertPivot(innerBatch, 20, "South");
        var added = commands.SetSlicerSelection(innerBatch, "FollowUpSlicer", ["North"], false);
        Assert.True(added.Success, added.ErrorMessage);
        Assert.Equal(["North", "South"], added.SelectedItems.Order());
        AssertPivot(innerBatch, 30, "North", "South");
        var cleared = commands.SetSlicerSelection(innerBatch, "FollowUpSlicer", []);
        Assert.True(cleared.Success, cleared.ErrorMessage);
        Assert.Equal(["North", "South"], cleared.SelectedItems.Order());
        AssertPivot(innerBatch, 30, "North", "South");
    }

    [Theory]
    [InlineData("IsSlicerCacheConnectedToPivot")]
    [InlineData("GetConnectedPivotTableNames")]
    public void ConnectedPivotScan_CancellationAfterEnteringHelper_StopsEnumeration(string helperName)
    {
        using var batch = ExcelSession.BeginBatch(fixture.CreateTestFile());
        var commands = new PivotTableCommands();
        PreparePivot(batch, commands);
        var created = commands.CreateSlicer(batch, "CancellationPivot", "Region",
            "ConnectedSlicer", "SlicerData", "E1");
        Assert.True(created.Success, created.ErrorMessage);
        Assert.Equal(["CancellationPivot"], created.ConnectedPivotTables);
        var filtered = commands.SetSlicerSelection(batch, "ConnectedSlicer", ["North"]);
        Assert.True(filtered.Success, filtered.ErrorMessage);
        Assert.Equal(["North"], filtered.SelectedItems);
        AssertPivot(batch, 10, "North");
        var helper = typeof(PivotTableCommands).GetMethod(
            helperName, BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(helper);
        using var cancellation = new CancellationTokenSource();
        Assert.False(cancellation.IsCancellationRequested);

        batch.Execute((context, _) =>
        {
            Excel.SlicerCaches? caches = null;
            Excel.SlicerCache? cache = null;
            Excel.SlicerPivotTables? pivots = null;
            Excel.PivotTable? pivot = null;
            try
            {
                caches = context.Book.SlicerCaches;
                cache = caches.Item[1];
                pivots = cache.PivotTables;
                Assert.Equal(1, pivots.Count);
                pivot = pivots.Item[1];
                var probe = new CancelOnConnectedCollectionRead(cache, cancellation);
                object?[] arguments = helperName == "IsSlicerCacheConnectedToPivot"
                    ? [probe, pivot, cancellation.Token]
                    : [probe, cancellation.Token];
                var error = Assert.Throws<TargetInvocationException>(() => helper.Invoke(null, arguments));
                var cancelled = Assert.IsType<OperationCanceledException>(error.InnerException);
                Assert.Equal(cancellation.Token, cancelled.CancellationToken);
                Assert.True(probe.ConnectedCollectionRead);
                Assert.True(cancellation.IsCancellationRequested);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
            }
        });
        Assert.Equal((1, 1), GetSlicerCounts(batch));
        var listed = commands.ListSlicers(batch, "CancellationPivot");
        Assert.True(listed.Success, listed.ErrorMessage);
        var slicer = Assert.Single(listed.Slicers);
        Assert.Equal(["CancellationPivot"], slicer.ConnectedPivotTables);
        Assert.Equal(["North"], slicer.SelectedItems);
        AssertPivot(batch, 10, "North");
        var replaced = commands.SetSlicerSelection(batch, "ConnectedSlicer", ["South"]);
        Assert.True(replaced.Success, replaced.ErrorMessage);
        Assert.Equal(["South"], replaced.SelectedItems);
        AssertPivot(batch, 20, "South");
        var followUp = commands.SetSlicerSelection(batch, "ConnectedSlicer", ["North"], false);
        Assert.True(followUp.Success, followUp.ErrorMessage);
        Assert.Equal(["North", "South"], followUp.SelectedItems.Order());
        AssertPivot(batch, 30, "North", "South");
    }

    // Forward real COM reads and inject cancellation only after the helper has started.
    private sealed class CancelOnConnectedCollectionRead(
        Excel.SlicerCache cache,
        CancellationTokenSource cancellation) : DynamicObject
    {
        public bool ConnectedCollectionRead { get; private set; }

        public override bool TryGetMember(GetMemberBinder binder, out object? result)
        {
            switch (binder.Name)
            {
                case "List":
                    result = cache.List;
                    return true;
                case "PivotTables":
                    result = cache.PivotTables;
                    ConnectedCollectionRead = true;
                    cancellation.Cancel();
                    return true;
                default:
                    result = null;
                    return false;
            }
        }
    }

    private static void PreparePivot(IExcelBatch batch, PivotTableCommands commands)
    {
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? data = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item[1];
                sheet.Name = "SlicerData";
                data = sheet.Range["A1:B3"];
                data.Value2 = new object[,] { { "Region", "Sales" }, { "North", 10 }, { "South", 20 } };
            }
            finally
            {
                ComUtilities.Release(ref data);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var created = commands.CreateFromRange(
            batch, "SlicerData", "A1:B3", "SlicerData", "H1", "CancellationPivot");
        Assert.True(created.Success, created.ErrorMessage);
        var field = commands.AddRowField(batch, "CancellationPivot", "Region");
        Assert.True(field.Success, field.ErrorMessage);
        var value = commands.AddValueField(batch, "CancellationPivot", "Sales");
        Assert.True(value.Success, value.ErrorMessage);
    }

    private static void AssertPivot(IExcelBatch batch, double total, params string[] regions) =>
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item["SlicerData"];
                range = sheet.Range["H1:I5"];
                var values = Assert.IsAssignableFrom<Array>((object)range.Value2);
                var rows = new List<(string Region, double Value)>();
                for (int row = values.GetLowerBound(0); row <= values.GetUpperBound(0); row++)
                {
                    if (values.GetValue(row, 2) is double value)
                        rows.Add((Convert.ToString(values.GetValue(row, 1),
                            System.Globalization.CultureInfo.InvariantCulture)!, value));
                }
                Assert.Equal(regions.Length + 1, rows.Count);
                Assert.Equal(regions.Order(), rows.Take(regions.Length).Select(row => row.Region).Order());
                foreach (var row in rows.Take(regions.Length))
                    Assert.Equal(row.Region == "North" ? 10d : 20d, row.Value);
                Assert.Equal(total, rows[^1].Value);
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private static (int Caches, int Visuals) GetSlicerCounts(IExcelBatch batch) =>
        batch.Execute((context, ct) =>
        {
            Excel.SlicerCaches? caches = null;
            Excel.SlicerCache? cache = null;
            Excel.Slicers? slicers = null;
            try
            {
                caches = context.Book.SlicerCaches;
                int visuals = 0;
                for (int index = 1; index <= caches.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        cache = caches.Item[index];
                        slicers = cache.Slicers;
                        visuals += slicers.Count;
                    }
                    finally
                    {
                        ComUtilities.Release(ref slicers);
                        ComUtilities.Release(ref cache);
                    }
                }
                return (caches.Count, visuals);
            }
            finally
            {
                ComUtilities.Release(ref caches);
            }
        });
}
