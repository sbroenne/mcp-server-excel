using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class PivotTableOlapDetectionTests
{
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
