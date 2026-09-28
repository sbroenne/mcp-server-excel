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
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var setResult = _pivotCommands.SetCacheOptions(
            batch,
            pivotName,
            refreshOnFileOpen: true,
            missingItemsLimit: PivotMissingItemsLimit.None,
            saveSourceData: false);
        Assert.True(setResult.Success, setResult.ErrorMessage);

        var getResult = _pivotCommands.GetCacheOptions(batch, pivotName);
        Assert.True(getResult.Success, getResult.ErrorMessage);
        Assert.False(getResult.IsOlap);
        Assert.True(getResult.RefreshOnFileOpen);
        Assert.Equal(
            PivotMissingItemsLimit.None,
            getResult.MissingItemsLimit);
        Assert.False(getResult.SaveSourceData);
    }
}
