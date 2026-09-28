using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePowerQueryMixedConnectionTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPowerQueryCommands _queries =
        fixture.CreateCommands<IPowerQueryCommands>();

    [Fact]
    public void Update_WithDataModelConnection_DoesNotThrowCOMException()
    {
        var dataModelQueryName = UniqueName("PQ_DM");
        var connectionOnlyQueryName = UniqueName("PQ_CO");
        CreateDataModelQuery(
            dataModelQueryName,
            "let Source = #table({\"ID\", \"Value\"}, {{1, 100}, {2, 200}}) in Source");
        CreateConnectionOnlyQuery(
            connectionOnlyQueryName,
            "let Source = #table({\"A\"}, {{1}}) in Source");

        var loadConfig = _queries.GetLoadConfig(
            _fixture.BatchToken,
            dataModelQueryName);
        Assert.True(loadConfig.Success, $"GetLoadConfig failed: {loadConfig.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfig.LoadMode);

        var updateResult = _queries.Update(
            _fixture.BatchToken,
            connectionOnlyQueryName,
            "let Source = #table({\"A\", \"B\"}, {{1, 2}}) in Source");

        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");
        var viewResult = _queries.View(_fixture.BatchToken, connectionOnlyQueryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("\"B\"", viewResult.MCode);
    }

    [Fact]
    public void View_WithDataModelConnection_DoesNotThrowCOMException()
    {
        var dataModelQueryName = UniqueName("PQ_DM");
        var connectionOnlyQueryName = UniqueName("PQ_CO");
        CreateDataModelQuery(
            dataModelQueryName,
            "let Source = #table({\"X\"}, {{1}}) in Source");
        CreateConnectionOnlyQuery(
            connectionOnlyQueryName,
            "let Source = #table({\"Y\"}, {{2}}) in Source");

        var viewResult = _queries.View(_fixture.BatchToken, connectionOnlyQueryName);

        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("\"Y\"", viewResult.MCode);
    }

    [Fact]
    public void List_WithDataModelConnection_Succeeds()
    {
        var dataModelQueryName = UniqueName("PQ_DM");
        var connectionOnlyQueryName = UniqueName("PQ_CO");
        CreateDataModelQuery(
            dataModelQueryName,
            "let Source = #table({\"X\"}, {{1}}) in Source");
        CreateConnectionOnlyQuery(
            connectionOnlyQueryName,
            "let Source = #table({\"Y\"}, {{2}}) in Source");

        var listResult = _queries.List(_fixture.BatchToken);

        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Queries);
        var queryNames = listResult.Queries.Select(query => query.Name).ToList();
        Assert.Contains(dataModelQueryName, queryNames);
        Assert.Contains(connectionOnlyQueryName, queryNames);
    }

    [Fact]
    public void Delete_WithDataModelConnection_Succeeds()
    {
        var dataModelQueryName = UniqueName("PQ_DM");
        var targetQueryName = UniqueName("PQ_Del");
        CreateDataModelQuery(
            dataModelQueryName,
            "let Source = #table({\"X\"}, {{1}}) in Source");
        CreateConnectionOnlyQuery(
            targetQueryName,
            "let Source = #table({\"Y\"}, {{2}}) in Source");

        var deleteResult = _queries.Delete(_fixture.BatchToken, targetQueryName);
        _fixture.ForgetPowerQuery(targetQueryName);

        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");
        var listResult = _queries.List(_fixture.BatchToken);
        Assert.True(listResult.Success);
        var queryNames = listResult.Queries.Select(query => query.Name).ToList();
        Assert.DoesNotContain(targetQueryName, queryNames);
        Assert.Contains(dataModelQueryName, queryNames);
    }

    [Fact]
    public void Update_WorksheetQuery_WithDataModelConnection_Succeeds()
    {
        var dataModelQueryName = UniqueName("PQ_DM");
        var worksheetQueryName = UniqueName("PQ_WS");
        CreateDataModelQuery(
            dataModelQueryName,
            "let Source = #table({\"X\"}, {{1}}) in Source");
        _queries.Create(
            _fixture.BatchToken,
            worksheetQueryName,
            "let Source = #table({\"Col1\"}, {{10}}) in Source",
            PowerQueryLoadMode.LoadToTable);
        _fixture.RegisterPowerQueryForCleanup(worksheetQueryName);
        _fixture.RegisterSheetForCleanup(worksheetQueryName);

        var updateResult = _queries.Update(
            _fixture.BatchToken,
            worksheetQueryName,
            "let Source = #table({\"Col1\", \"Col2\"}, {{10, 20}}) in Source");

        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");
        var viewResult = _queries.View(_fixture.BatchToken, worksheetQueryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("\"Col2\"", viewResult.MCode);
    }

    [Fact]
    public void DataModelLoad_CreatesNonOledbConnection()
    {
        var queryName = UniqueName("PQ_Verify");
        CreateDataModelQuery(
            queryName,
            "let Source = #table({\"V\"}, {{1}}) in Source");

        var hasNonOledbConnection = _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? connections = null;
            try
            {
                connections = ctx.Book.Connections;
                var count = (int)connections.Count;
                for (var index = 1; index <= count; index++)
                {
                    dynamic? connection = null;
                    try
                    {
                        connection = connections[index];
                        var connectionType = (int)connection.Type;
                        if (connectionType != 1)
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref connection);
                    }
                }

                return false;
            }
            finally
            {
                ComUtilities.Release(ref connections);
            }
        });

        Assert.True(
            hasNonOledbConnection,
            "Expected at least one non-OLEDB connection (Type != 1) after LoadToDataModel.");
    }

    private void CreateDataModelQuery(string queryName, string mCode)
    {
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            mCode,
            PowerQueryLoadMode.LoadToDataModel);
        _fixture.RegisterPowerQueryForCleanup(queryName);
    }

    private void CreateConnectionOnlyQuery(string queryName, string mCode)
    {
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            mCode,
            PowerQueryLoadMode.ConnectionOnly);
        _fixture.RegisterPowerQueryForCleanup(queryName);
    }

    private static string UniqueName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..Math.Min(prefix.Length + 9, 31)];
}
