using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Slicer checks that must happen before Excel creates a slicer cache.
/// </summary>
public sealed partial class PersistentServicePivotTableTests
{
    [Theory]
    [InlineData("NoSuchSheet", "I2", "NoSuchSheet")]
    [InlineData(null, "NotACell", "NotACell")]
    public void CreateSlicer_BadDestination_DoesNotLeaveSlicerCache(string? sheet, string position, string expectedInMessage)
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SlicerDestinationPivot"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "SlicerDestinationPivot", "Region"));

        var error = Assert.ThrowsAny<Exception>(() => _pivotCommands.CreateSlicer(
            batch, "SlicerDestinationPivot", "Region", "BadDestinationSlicer", sheet ?? _salesSheetName, position));

        Assert.Contains(expectedInMessage, error.Message, StringComparison.Ordinal);
        Assert.Equal(0, CountSlicerCaches());
        AssertOriginalSales();
    }

    [Fact]
    public void CreateSlicer_NameAlreadyUsed_DoesNotLeaveSlicerCache()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SlicerNamePivot"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "SlicerNamePivot", "Region"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "SlicerNamePivot", "Product"));
        RequireSuccess(_pivotCommands.CreateSlicer(
            batch, "SlicerNamePivot", "Region", "TakenSlicerName", _salesSheetName, "I2"));

        var error = Assert.ThrowsAny<Exception>(() => _pivotCommands.CreateSlicer(
            batch, "SlicerNamePivot", "Product", "TakenSlicerName", _salesSheetName, "I12"));

        Assert.Contains("TakenSlicerName", error.Message, StringComparison.Ordinal);
        Assert.Contains("already", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, CountSlicerCaches());
        AssertPivotSlicer("TakenSlicerName", "SlicerNamePivot", "Region", "I2", ["North", "South"]);
        AssertOriginalSales();
    }

    [Fact]
    public void CreateSlicer_DestinationSheetInDifferentCase_UsesExistingSheet()
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SlicerCasePivot"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "SlicerCasePivot", "Region"));

        RequireSuccess(_pivotCommands.CreateSlicer(
            batch, "SlicerCasePivot", "Region", "CaseSlicer", _salesSheetName.ToUpperInvariant(), "I2"));

        AssertPivotSlicer("CaseSlicer", "SlicerCasePivot", "Region", "I2", ["North", "South"]);
        AssertOriginalSales();
    }

    private int CountSlicerCaches() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Microsoft.Office.Interop.Excel.SlicerCaches? caches = null;
            try
            {
                caches = context.Book.SlicerCaches;
                return caches.Count;
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref caches);
            }
        });
}
