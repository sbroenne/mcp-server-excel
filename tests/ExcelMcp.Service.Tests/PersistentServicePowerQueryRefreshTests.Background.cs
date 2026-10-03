using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryRefreshTests
{
    [Fact]
    public void Refresh_DataModelQuery_LoadsChangedDataWithoutWorksheetDestination()
    {
        var queryName = UniqueName("DM_BGQ");
        var created = RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            ValidMCode,
            PowerQueryLoadMode.LoadToDataModel));
        Assert.True(created.Success, created.ErrorMessage);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _storedSources.Add(queryName, ValidMCode);
        AssertModelValue(queryName, 1);
        StageSource(queryName, "let Source = #table({\"X\"}, {{87}}) in Source");
        AssertModelValue(queryName, 1);

        var result = RequireSuccess(_queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromMinutes(2)));

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        AssertRefreshMetadata(result, queryName, null);
        AssertModelValue(queryName, 87);
    }

    [Fact]
    public void Refresh_WorksheetQuery_CompletesSuccessfully()
    {
        var queryName = CreateWorksheetQuery("WS_BGQ");
        StageWorksheetUpdate(queryName, 93);

        var result = RequireSuccess(_queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromMinutes(2)));

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.Equal(queryName, result.LoadedToSheet);
        AssertRefreshMetadata(result, queryName, queryName);
        AssertWorksheetValue(queryName, 93);
    }
}
