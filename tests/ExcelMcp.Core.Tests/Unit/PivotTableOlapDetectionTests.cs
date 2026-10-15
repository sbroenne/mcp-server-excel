using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PivotTable;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelProxy = Sbroenne.ExcelMcp.Core.Tests.Unit.RangeHelpersExceptionTests.ExcelProxy;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class PivotTableOlapDetectionTests
{
    [Fact]
    public void DisconnectProxy_ReachesCubeFieldsThroughTypedExcelLookup()
    {
        var owner = CreateFailingPivot(ResiliencePipelines.RPC_E_DISCONNECTED);
        var typedPivot = CreatePivot(owner);
        var typedPivots = ExcelProxy.Create<Excel.PivotTables>((member, args) =>
        {
            Assert.Equal("Item", member);
            Assert.Equal("Pivot", args![0]);
            return typedPivot;
        });
        var typedSheet = ExcelProxy.Create<Excel.Worksheet>((member, _) =>
        {
            Assert.Equal("PivotTables", member);
            return typedPivots;
        });
        var sheets = ExcelProxy.Create<Excel.Sheets>((member, args) =>
        {
            Assert.Equal("get__Default", member);
            Assert.Equal("Sheet", args![0]);
            return typedSheet;
        });
        var sheet = (Excel.Worksheet)sheets["Sheet"];
        var pivots = (Excel.PivotTables)sheet.PivotTables();
        var pivot = pivots.Item("Pivot");

        var error = Assert.Throws<COMException>(() => PivotTableHelpers.IsOlapPivotTable(pivot));

        Assert.Same(owner.Failure, error);
        Assert.Equal(ResiliencePipelines.RPC_E_DISCONNECTED, error.HResult);
        Assert.True(owner.CubeFieldsReadAttempted);
    }

    [Theory]
    [InlineData(ResiliencePipelines.RPC_E_CALL_FAILED)]
    [InlineData(ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE)]
    [InlineData(ResiliencePipelines.RPC_E_DISCONNECTED)]
    public void IsOlapPivotTable_DisconnectedCubeFields_PropagatesOriginalFailure(int hresult)
    {
        var pivot = CreateFailingPivot(hresult);
        var typedPivot = CreatePivot(pivot);

        var error = Assert.Throws<COMException>(() => PivotTableHelpers.IsOlapPivotTable(typedPivot));

        Assert.Same(pivot.Failure, error);
        Assert.Equal(hresult, error.HResult);
        Assert.True(pivot.CubeFieldsReadAttempted);
    }

    [Fact]
    public void IsOlapPivotTable_UnavailableCubeFields_RemainsNonOlap()
    {
        var pivot = CreateFailingPivot(unchecked((int)0x800A03EC));
        var typedPivot = CreatePivot(pivot);

        Assert.False(PivotTableHelpers.IsOlapPivotTable(typedPivot));
        Assert.True(pivot.CubeFieldsReadAttempted);
    }

    private static Excel.PivotTable CreatePivot(ReportFilterCancellationTests.DisconnectedBatch owner) =>
        ExcelProxy.Create<Excel.PivotTable>((member, _) =>
        {
            Assert.Equal("get_CubeFields", member);
            owner.CubeFieldsReadAttempted = true;
            throw owner.Failure;
        });

    private static ReportFilterCancellationTests.DisconnectedBatch CreateFailingPivot(int hresult)
    {
#pragma warning disable CA2201 // Simulate Excel's CubeFields failure without starting Excel.
        return new ReportFilterCancellationTests.DisconnectedBatch
        {
            Failure = new COMException("Injected CubeFields failure.", hresult)
        };
#pragma warning restore CA2201
    }
}
