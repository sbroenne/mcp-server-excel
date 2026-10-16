using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PivotTable;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class PivotTableOlapDetectionTests
{
    [Theory]
    [InlineData(-2146826246, true)]
    [InlineData(-2146827284, false)]
    [InlineData(-2147417848, false)]
    [InlineData(123, false)]
    [InlineData("Values", false)]
    [InlineData("Sales", false)]
    [InlineData(null, false)]
    public void ValuesPlaceholder_RequiresExactNativeSourceNameSentinel(object? sourceName, bool expected) =>
        Assert.Equal(expected, PivotTableHelpers.IsValuesPlaceholderSourceName(sourceName));

    [Fact]
    public void DisconnectProxy_ReachesCubeFieldsThroughTypedExcelLookup()
    {
        var owner = new ReportFilterCancellationTests.DisconnectedBatch
        {
            Failure = new FailingPivot(ResiliencePipelines.RPC_E_DISCONNECTED).Failure
        };
        var sheets = ReportFilterCancellationTests.DisconnectedPivotObjects.Create<Excel.Sheets>(owner);
        var sheet = (Excel.Worksheet)sheets["Sheet"];
        var pivots = (Excel.PivotTables)sheet.PivotTables();
        var pivot = pivots.Item("Pivot");

        Assert.Same(owner.Failure,
            Assert.Throws<COMException>(() => PivotTableHelpers.IsOlapPivotTable(pivot)));
        Assert.True(owner.CubeFieldsReadAttempted);
    }

    [Theory]
    [InlineData(ResiliencePipelines.RPC_E_CALL_FAILED)]
    [InlineData(ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE)]
    [InlineData(ResiliencePipelines.RPC_E_DISCONNECTED)]
    public void IsOlapPivotTable_DisconnectedCubeFields_PropagatesOriginalFailure(int hresult)
    {
        var pivot = new FailingPivot(hresult);

        var error = Assert.Throws<COMException>(() => PivotTableHelpers.IsOlapPivotTable(pivot));

        Assert.Same(pivot.Failure, error);
        Assert.True(pivot.CubeFieldsReadAttempted);
    }

    [Fact]
    public void IsOlapPivotTable_AmbiguousNativeFailure_PropagatesOriginalError()
    {
        var pivot = new FailingPivot(unchecked((int)0x800A03EC));

        Assert.Same(pivot.Failure, Assert.Throws<COMException>(() => PivotTableHelpers.IsOlapPivotTable(pivot)));
        Assert.True(pivot.CubeFieldsReadAttempted);
    }

    [Fact]
    public void RegularPivot_DoesNotProbeOlapOnlyCubeFields()
    {
        var pivot = new FailingPivot(unchecked((int)0x800A03EC), false);

        Assert.False(PivotTableHelpers.IsOlapPivotTable(pivot));
        Assert.Equal(1, pivot.CacheReadCount);
        Assert.False(pivot.CubeFieldsReadAttempted);
    }

    [Fact]
    public void TryGetCubeFields_RegularCache_DoesNotProbeOlapOnlyMembers()
    {
        var pivot = new FailingPivot(unchecked((int)0x800A03EC), false);

        Assert.False(PivotTableHelpers.TryGetCubeFields(pivot, out dynamic? fields));

        Assert.Null((object?)fields);
        Assert.Equal(1, pivot.CacheReadCount);
        Assert.False(pivot.CubeFieldsReadAttempted);
    }

    [Fact]
    public void TryGetCubeFields_AmbiguousNativeFailure_PropagatesOriginalError()
    {
        var pivot = new FailingPivot(unchecked((int)0x800A03EC));

        var error = Assert.Throws<COMException>(() => PivotTableHelpers.TryGetCubeFields(pivot, out _));

        Assert.Same(pivot.Failure, error);
        Assert.Equal(1, pivot.CacheReadCount);
        Assert.True(pivot.CubeFieldsReadAttempted);
    }

    [Fact]
    public void IsOlapPivotTable_CacheMethodFailure_PropagatesOriginalError()
    {
        var pivot = new FailingCachePivot();

        Assert.Same(pivot.Failure, Assert.Throws<COMException>(() => PivotTableHelpers.IsOlapPivotTable(pivot)));
    }

    public sealed class FailingPivot(int hresult, bool olap = true)
    {
        private readonly CacheState _cache = new(olap);
        public int CacheReadCount { get; private set; }

        public CacheState PivotCache()
        {
            CacheReadCount++;
            return _cache;
        }

#pragma warning disable CA2201 // Simulate Excel's CubeFields failure without starting Excel.
        public COMException Failure { get; } = new("Injected CubeFields failure.", hresult);
#pragma warning restore CA2201
        public bool CubeFieldsReadAttempted { get; private set; }

        public object CubeFields
        {
            get
            {
                CubeFieldsReadAttempted = true;
                throw Failure;
            }
        }
    }

    public sealed class FailingCachePivot
    {
#pragma warning disable CA2201 // Simulate a native PivotCache() acquisition failure without starting Excel.
        public COMException Failure { get; } = new("Injected PivotCache failure.", unchecked((int)0x800A03EC));
#pragma warning restore CA2201
        public CacheState PivotCache() => throw Failure;
    }

    public sealed class CacheState(bool olap)
    {
        public bool OLAP { get; } = olap;
    }
}
