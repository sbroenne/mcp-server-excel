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
        var (_, owner) = CreatePivot(ResiliencePipelines.RPC_E_DISCONNECTED);
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
        var (pivot, owner) = CreatePivot(hresult);

        var error = Assert.Throws<COMException>(() => PivotTableHelpers.IsOlapPivotTable(pivot));

        Assert.Same(owner.Failure, error);
        Assert.True(owner.CubeFieldsReadAttempted);
    }

    [Fact]
    public void IsOlapPivotTable_AmbiguousNativeFailure_PropagatesOriginalError()
    {
        var (pivot, owner) = CreatePivot(unchecked((int)0x800A03EC));

        Assert.Same(owner.Failure, Assert.Throws<COMException>(() => PivotTableHelpers.IsOlapPivotTable(pivot)));
        Assert.True(owner.CubeFieldsReadAttempted);
    }

    [Fact]
    public void RegularPivot_DoesNotProbeOlapOnlyCubeFields()
    {
        var (pivot, owner) = CreatePivot(unchecked((int)0x800A03EC), olap: false);

        Assert.False(PivotTableHelpers.IsOlapPivotTable(pivot));
        Assert.Equal(1, owner.CacheReadCount);
        Assert.False(owner.CubeFieldsReadAttempted);
    }

    [Fact]
    public void TryGetCubeFields_RegularCache_DoesNotProbeOlapOnlyMembers()
    {
        var (pivot, owner) = CreatePivot(unchecked((int)0x800A03EC), olap: false);

        Assert.False(PivotTableHelpers.TryGetCubeFields(pivot, out Excel.CubeFields? fields));

        Assert.Null(fields);
        Assert.Equal(1, owner.CacheReadCount);
        Assert.False(owner.CubeFieldsReadAttempted);
    }

    [Fact]
    public void TryGetCubeFields_AmbiguousNativeFailure_PropagatesOriginalError()
    {
        var (pivot, owner) = CreatePivot(unchecked((int)0x800A03EC));

        var error = Assert.Throws<COMException>(() => PivotTableHelpers.TryGetCubeFields(pivot, out _));

        Assert.Same(owner.Failure, error);
        Assert.Equal(1, owner.CacheReadCount);
        Assert.True(owner.CubeFieldsReadAttempted);
    }

    [Fact]
    public void IsOlapPivotTable_CacheMethodFailure_PropagatesOriginalError()
    {
        var (pivot, owner) = CreatePivot(unchecked((int)0x800A03EC), failCacheRead: true);

        Assert.Same(owner.Failure, Assert.Throws<COMException>(() => PivotTableHelpers.IsOlapPivotTable(pivot)));
        Assert.Equal(1, owner.CacheReadCount);
        Assert.False(owner.CubeFieldsReadAttempted);
    }

    private static (Excel.PivotTable Pivot, ReportFilterCancellationTests.DisconnectedBatch Owner) CreatePivot(
        int hresult, bool olap = true, bool failCacheRead = false)
    {
        var owner = new ReportFilterCancellationTests.DisconnectedBatch
        {
#pragma warning disable CA2201 // Inject a native PivotTable failure without starting Excel.
            Failure = new COMException("Injected PivotTable failure.", hresult),
#pragma warning restore CA2201
            Olap = olap,
            FailCacheRead = failCacheRead
        };
        return (ReportFilterCancellationTests.DisconnectedPivotObjects.Create<Excel.PivotTable>(owner), owner);
    }
}
