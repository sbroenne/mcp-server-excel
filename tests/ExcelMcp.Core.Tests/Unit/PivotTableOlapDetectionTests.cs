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
    public void IsOlapPivotTable_UnavailableCubeFields_RemainsNonOlap()
    {
        var pivot = new FailingPivot(unchecked((int)0x800A03EC));

        Assert.False(PivotTableHelpers.IsOlapPivotTable(pivot));
        Assert.True(pivot.CubeFieldsReadAttempted);
    }

    public sealed class FailingPivot(int hresult)
    {
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
}
