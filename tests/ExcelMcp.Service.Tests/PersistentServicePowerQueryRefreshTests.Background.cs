using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryRefreshTests
{
    [Fact]
    public void Refresh_DataModelQuery_CompletesWithoutCpuSpin()
    {
        var queryName = UniqueName("DM_BGQ");
        const string mCode = """
            let
                Source = #table(
                    {"ID", "Name", "Value"},
                    {
                        {1, "Alpha", 100},
                        {2, "Beta", 200},
                        {3, "Gamma", 300}
                    })
            in
                Source
            """;
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            mCode,
            PowerQueryLoadMode.LoadToDataModel);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var result = _queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromMinutes(2));

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.Equal(queryName, result.QueryName);
        Assert.True(
            result.IsConnectionOnly || string.IsNullOrEmpty(result.LoadedToSheet),
            "Data Model query should not be loaded to a worksheet.");
    }

    [Fact]
    public void Refresh_WorksheetQuery_CompletesSuccessfully()
    {
        var queryName = UniqueName("WS_BGQ");
        const string mCode = """
            let
                Source = #table(
                    {"ID", "Name"},
                    {{1, "Alpha"}, {2, "Beta"}})
            in
                Source
            """;
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            mCode,
            PowerQueryLoadMode.LoadToTable,
            queryName);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(queryName);

        var result = _queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromMinutes(2));

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(
            string.IsNullOrEmpty(result.LoadedToSheet),
            "Worksheet query should have a loaded sheet.");
    }
}
