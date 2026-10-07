using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Table slicer checks that must happen before Excel creates a slicer cache.
/// </summary>
public sealed partial class PersistentServiceTablePreflightTests
{
    [Theory]
    [InlineData("NoSuchSheet", "F2", "NoSuchSheet")]
    [InlineData("Sales", "NotACell", "NotACell")]
    public void CreateTableSlicer_BadDestination_DoesNotLeaveSlicerCache(string sheet, string position, string expectedInMessage)
    {
        var batch = _fixture.BatchToken;

        var error = Assert.ThrowsAny<Exception>(() => _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "BadDestinationSlicer", sheet, position));

        Assert.Contains(expectedInMessage, error.Message, StringComparison.Ordinal);
        Assert.Equal(0, CountTableSlicerCaches());
        AssertVisibleRegions(["North", "South", "East", "West"]);
    }

    [Fact]
    public void CreateTableSlicer_NameAlreadyUsed_DoesNotLeaveSlicerCache()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "TakenSlicerName", "Sales", "F2"));

        var error = Assert.ThrowsAny<Exception>(() => _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Product", "TakenSlicerName", "Sales", "F12"));

        Assert.Contains("TakenSlicerName", error.Message, StringComparison.Ordinal);
        Assert.Contains("already", error.Message, StringComparison.OrdinalIgnoreCase);
        AssertNativeSingleSlicer("TakenSlicerName", "Region", "SalesTable", "Sales", "F2");
        AssertVisibleRegions(["North", "South", "East", "West"]);
    }

    [Fact]
    public void CreateTableSlicer_DestinationSheetInDifferentCase_UsesExistingSheet()
    {
        var batch = _fixture.BatchToken;

        RequireSuccess(_tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "CaseSlicer", "SALES", "F2"));

        AssertNativeSingleSlicer("CaseSlicer", "Region", "SalesTable", "Sales", "F2");
        AssertVisibleRegions(["North", "South", "East", "West"]);
    }

    private int CountTableSlicerCaches() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.SlicerCaches? caches = null;
            try
            {
                caches = context.Book.SlicerCaches;
                return caches.Count;
            }
            finally
            {
                ComUtilities.Release(ref caches);
            }
        });
}
