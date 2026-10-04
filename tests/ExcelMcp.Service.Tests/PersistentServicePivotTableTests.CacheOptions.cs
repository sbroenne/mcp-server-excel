using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    public void CacheOptions_RegularCache_RoundTripsMutableSettings()
    {
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);
        var pivotName = $"Cache_{Guid.NewGuid():N}";
        var createResult = _pivotCommands.CreateFromRange(
            batch,
            _salesSheetName,
            "A1:D6",
            destinationSheet,
            "A1",
            pivotName);
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));
        AssertPivotSales(325, 325, pivotName);
        var before = ReadNativeCacheOptions(destinationSheet, pivotName);

        var setResult = _pivotCommands.SetCacheOptions(
            batch,
            pivotName,
            refreshOnFileOpen: true,
            missingItemsLimit: PivotMissingItemsLimit.None,
            saveSourceData: false);
        RequireSuccess(setResult);

        var getResult = _pivotCommands.GetCacheOptions(batch, pivotName);
        RequireSuccess(getResult);
        Assert.False(getResult.IsOlap);
        Assert.True(getResult.RefreshOnFileOpen);
        Assert.Equal(
            PivotMissingItemsLimit.None,
            getResult.MissingItemsLimit);
        Assert.False(getResult.SaveSourceData);
        RequireSuccess(getResult);
        var expected = before with { RefreshOnOpen = true, MissingItems = 0, SaveData = false };
        Assert.Equal(expected, ReadNativeCacheOptions(destinationSheet, pivotName));
        Assert.Equal(before.EnableRefresh, getResult.EnableRefresh);
        Assert.Equal(before.Optimize, getResult.OptimizeCache);
        AssertPivotSales(325, 325, pivotName);
        AssertOriginalSales();

        var reset = RequireSuccess(_pivotCommands.SetCacheOptions(
            batch, pivotName, refreshOnFileOpen: false, saveSourceData: true));
        Assert.False(reset.RefreshOnFileOpen);
        Assert.True(reset.SaveSourceData);
        Assert.Equal(expected with { RefreshOnOpen = false, SaveData = true },
            ReadNativeCacheOptions(destinationSheet, pivotName));
        AssertPivotSales(325, 325, pivotName);
        AssertOriginalSales();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CacheOptions_MissingPivot_PreservesConfiguredCache(bool write)
    {
        var batch = _fixture.BatchToken;
        RequireSuccess(_pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "RetainedCache"));
        RequireSuccess(_pivotCommands.AddRowField(batch, "RetainedCache", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "RetainedCache", "Sales"));
        RequireSuccess(_pivotCommands.SetCacheOptions(
            batch, "RetainedCache", refreshOnFileOpen: true, missingItemsLimit: PivotMissingItemsLimit.None));
        var before = ReadNativeCacheOptions(_salesSheetName, "RetainedCache");

        var error = Assert.Throws<InvalidOperationException>(() =>
        {
            if (write)
            {
                _pivotCommands.SetCacheOptions(batch, "MissingCache", refreshOnFileOpen: false);
            }
            else
            {
                _pivotCommands.GetCacheOptions(batch, "MissingCache");
            }
        });

        Assert.Contains("not found", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, ReadNativeCacheOptions(_salesSheetName, "RetainedCache"));
        AssertPivotSales(325, 325, "RetainedCache");
        AssertOriginalSales();
    }

    private sealed record NativeCacheOptions(
        bool EnableRefresh, bool RefreshOnOpen, bool Optimize, bool SaveData, int MissingItems);

    private NativeCacheOptions ReadNativeCacheOptions(string sheetName, string pivotName) =>
        ReadNativePivot(sheetName, pivotName, pivot =>
        {
            Microsoft.Office.Interop.Excel.PivotCache? cache = null;
            try
            {
                cache = pivot.PivotCache();
                Assert.False(cache.OLAP);
                return new NativeCacheOptions(cache.EnableRefresh, cache.RefreshOnFileOpen,
                    cache.OptimizeCache, pivot.SaveData, (int)cache.MissingItemsLimit);
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref cache);
            }
        });
}
