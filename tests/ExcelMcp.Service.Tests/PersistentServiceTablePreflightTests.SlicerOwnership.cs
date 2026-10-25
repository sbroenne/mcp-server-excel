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
                var probeSlicers = SlicerOwnershipProxy.Create<Excel.Slicers>((member, args) =>
                    member.Name switch
                    {
                        "get_Count" => probe.Count,
                        "get_Item" or "get__Default" or "Item" => probe.Item(Assert.IsType<int>(args![0])),
                        _ => throw new NotSupportedException(member.Name)
                    });
                var probeCache = SlicerOwnershipProxy.Create<Excel.SlicerCache>((member, _) =>
                    member.Name switch
                    {
                        "get_List" => true,
                        "get_Slicers" => probeSlicers,
                        _ => throw new NotSupportedException(member.Name)
                    });
                var probeCaches = SlicerOwnershipProxy.Create<Excel.SlicerCaches>((member, args) =>
                {
                    if (member.Name == "get_Count")
                        return 1;
                    if (member.Name is "get_Item" or "get__Default" or "Item")
                    {
                        Assert.Equal(1, Assert.IsType<int>(args![0]));
                        return probeCache;
                    }
                    throw new NotSupportedException(member.Name);
                });

                var wrapper = Assert.Throws<TargetInvocationException>(() =>
                    lookup.Invoke(null, [probeCaches, "OwnershipSlicer"]));
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

    public class SlicerOwnershipProxy : DispatchProxy
    {
        public Func<MethodInfo, object?[]?, object?> Read { get; set; } = null!;

        public static T Create<T>(Func<MethodInfo, object?[]?, object?> read) where T : class
        {
            var interfaceType = typeof(TableCommands).Assembly.GetType(typeof(T).FullName!, throwOnError: true)!;
            var value = (T)Create(interfaceType, typeof(SlicerOwnershipProxy));
            ((SlicerOwnershipProxy)(object)value).Read = read;
            return value;
        }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args) =>
            Read(targetMethod ?? throw new InvalidOperationException("No intercepted slicer interface member."), args);
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
