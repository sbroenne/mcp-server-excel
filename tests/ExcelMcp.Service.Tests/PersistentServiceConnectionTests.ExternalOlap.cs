using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapSchema_UsesSelectedCubeAndContinuesThroughService()
    {
        string connectionString = Environment.GetEnvironmentVariable(
            "EXCELMCP_TEST_OLAP_CONNECTION_STRING")!;
        string cubeName = Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_CUBE")!;
        string hierarchyUniqueName = Environment.GetEnvironmentVariable(
            "EXCELMCP_TEST_OLAP_HIERARCHY")!;
        string levelUniqueName = Environment.GetEnvironmentVariable(
            "EXCELMCP_TEST_OLAP_LEVEL")!;
        var connectionName = UniqueConnectionName("ExternalOlap");

        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            try
            {
                connections = context.Book.Connections;
                connection = connections.Add2(
                    connectionName,
                    "Configured external OLAP integration fixture",
                    connectionString,
                    cubeName,
                    Excel.XlCmdType.xlCmdCube,
                    false,
                    false);
                Assert.Equal(connectionName, connection.Name);
                _fixture.RegisterConnectionForCleanup(connectionName);
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref connection);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref connections);
            }
        });

        var schema = RequireSuccess(_connections.DiscoverOlapSchema(
            _fixture.BatchToken,
            connectionName,
            hierarchyUniqueName));
        Assert.Contains(schema.Hierarchies, hierarchy => hierarchy.UniqueName == hierarchyUniqueName);
        Assert.Contains(schema.Levels, level => level.UniqueName == levelUniqueName);
        Assert.DoesNotContain(schema.Hierarchies, hierarchy =>
            hierarchy.UniqueName != hierarchyUniqueName
            && schema.Levels.Any(level => level.HierarchyUniqueName == hierarchy.UniqueName));

        var firstPage = RequireSuccess(_connections.SearchOlapMembers(
            _fixture.BatchToken,
            connectionName,
            hierarchyUniqueName,
            levelUniqueName,
            pageSize: 1));
        var firstMember = Assert.Single(firstPage.Members);
        Assert.Equal(1, firstPage.ReturnedCount);
        Assert.False(string.IsNullOrWhiteSpace(firstPage.ContinuationToken));

        var secondPage = RequireSuccess(_connections.SearchOlapMembers(
            _fixture.BatchToken,
            connectionName,
            hierarchyUniqueName,
            levelUniqueName,
            continuationToken: firstPage.ContinuationToken,
            pageSize: 1));
        var secondMember = Assert.Single(secondPage.Members);
        Assert.NotEqual(firstMember.UniqueName, secondMember.UniqueName);
    }
}

[AttributeUsage(AttributeTargets.Method)]
public sealed class ConfiguredExternalOlapFactAttribute : FactAttribute
{
    private static readonly string[] RequiredSettings =
    [
        "EXCELMCP_TEST_OLAP_CONNECTION_STRING",
        "EXCELMCP_TEST_OLAP_CUBE",
        "EXCELMCP_TEST_OLAP_HIERARCHY",
        "EXCELMCP_TEST_OLAP_LEVEL"
    ];

    public ConfiguredExternalOlapFactAttribute()
    {
        if (RequiredSettings.Any(name => string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(name))))
        {
            Skip = "Configure an external OLAP test cube and its hierarchy/level environment variables.";
        }
    }
}
