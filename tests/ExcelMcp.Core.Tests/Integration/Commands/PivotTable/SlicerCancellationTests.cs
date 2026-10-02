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
    [Fact]
    public void CreateSlicer_CancelledCallbackWithNoCache_StopsBeforeCreation()
    {
        using var innerBatch = ExcelSession.BeginBatch(fixture.CreateTestFile());
        var commands = new PivotTableCommands();
        PreparePivot(innerBatch, commands);
        var initialCounts = GetSlicerCounts(innerBatch);
        Assert.Equal((0, 0), initialCounts);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var batch = new InjectedCancellationBatch(innerBatch, cancellation.Token);

        Assert.ThrowsAny<OperationCanceledException>(() =>
            commands.CreateSlicer(batch, "CancellationPivot", "Region",
                "CancelledSlicer", "SlicerData", "E1"));
        Assert.Equal(initialCounts, GetSlicerCounts(innerBatch));

        var followUp = commands.CreateSlicer(innerBatch, "CancellationPivot", "Region",
            "FollowUpSlicer", "SlicerData", "E1");
        Assert.True(followUp.Success, followUp.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(followUp.ErrorMessage));
        Assert.Equal(["North", "South"], followUp.AvailableItems.Order());
        Assert.Equal(["North", "South"], followUp.SelectedItems.Order());
        Assert.Equal((1, 1), GetSlicerCounts(innerBatch));
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
        Assert.Equal(["CancellationPivot"], Assert.Single(listed.Slicers).ConnectedPivotTables);
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
    }

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
