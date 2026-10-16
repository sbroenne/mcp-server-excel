using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceTablePreflightTests
{
    [Fact]
    public void TableSlicerLookup_FailedNativeNameRead_ReleasesAcquiredReference()
    {
        RequireSuccess(_tableCommands.CreateTableSlicer(
            _fixture.BatchToken, "SalesTable", "Region", "OwnershipSlicer", "Sales", "F2"));

        // Service cannot inject a failure between this internal helper's acquisition and name read.
        _fixture.ExecuteRawVerification((context, cancellationToken) =>
        {
            Excel.SlicerCaches? caches = null;
            Excel.SlicerCache? cache = null;
            Excel.Slicers? slicers = null;
            DeletedNativeSlicerProbe? probe = null;
            try
            {
                caches = context.Book.SlicerCaches;
                cache = caches.Item[1];
                slicers = cache.Slicers;
                Assert.Equal(1, slicers.Count);
                var lookup = typeof(TableCommands).GetMethod("FindTableSlicerByName",
                    BindingFlags.Static | BindingFlags.NonPublic);
                Assert.NotNull(lookup);
                probe = new DeletedNativeSlicerProbe(slicers);

                var wrapper = Assert.Throws<TargetInvocationException>(() =>
                    lookup.Invoke(null, [new SlicerOwnershipCaches(probe), "OwnershipSlicer"]));
                Assert.False(ComErrors.IsMemberUnavailable(Assert.IsType<COMException>(wrapper.InnerException)));
                Assert.True(probe.Deleted);

                Assert.NotNull(probe.Acquired);
                Assert.Throws<InvalidComObjectException>(() => _ = probe.Acquired.Name);
            }
            finally
            {
                var acquired = probe?.Acquired;
                ComUtilities.Release(ref acquired);
                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
            }
            return 0;
        });
        Assert.Empty(RequireSuccess(_tableCommands.ListTableSlicers(_fixture.BatchToken)).Slicers);
        AssertVisibleRegions(["North", "South", "East", "West"]);
    }

    public sealed class SlicerOwnershipCaches(DeletedNativeSlicerProbe slicers)
    {
        public int Count { get; } = 1;
        public SlicerOwnershipCache Item(int index)
        {
            Assert.Equal(1, index);
            return new SlicerOwnershipCache(slicers);
        }
    }

    public sealed class SlicerOwnershipCache(DeletedNativeSlicerProbe slicers)
    {
        public bool List { get; } = true;
        public DeletedNativeSlicerProbe Slicers => slicers;
    }

    public sealed class DeletedNativeSlicerProbe(Excel.Slicers slicers)
    {
        public int Count => slicers.Count;
        public bool Deleted { get; private set; }
        public Excel.Slicer? Acquired { get; private set; }

        public Excel.Slicer Item(int index)
        {
            Excel.Slicer? acquired = slicers.Item[index];
            try
            {
                Assert.Equal("OwnershipSlicer", acquired.Name);
                Acquired = acquired;
                acquired.Delete();
                Deleted = true;
                var result = acquired;
                acquired = null;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref acquired);
            }
        }
    }
}
