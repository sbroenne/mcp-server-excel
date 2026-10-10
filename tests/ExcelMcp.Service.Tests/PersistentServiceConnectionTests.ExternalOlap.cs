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
        var pivotSheetName = _fixture.CreateTestSheet(_fixture.BatchToken);

        try
        {
            CreateOlapPivotTable(connectionString, cubeName, connectionName, pivotSheetName);
            AssertSchemaAndMemberPaging(connectionName, hierarchyUniqueName, levelUniqueName);
        }
        finally
        {
            _fixture.Send("sheet.delete", new { sheetName = pivotSheetName });
            _fixture.ForgetSheet(pivotSheetName);
        }
    }

    // Excel only keeps an OLAP session open while a PivotTable uses the connection,
    // which is how users reach external cubes in practice.
    private void CreateOlapPivotTable(
        string connectionString,
        string cubeName,
        string connectionName,
        string pivotSheetName)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.OLEDBConnection? oledbConnection = null;
            Excel.PivotCaches? pivotCaches = null;
            Excel.PivotCache? pivotCache = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? destination = null;
            Excel.PivotTable? pivotTable = null;
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
                oledbConnection = connection.OLEDBConnection;
                oledbConnection.MaintainConnection = true;

                pivotCaches = context.Book.PivotCaches();
                pivotCache = pivotCaches.Create(
                    Excel.XlPivotTableSourceType.xlExternal,
                    connection,
                    Excel.XlPivotTableVersionList.xlPivotTableVersion15);
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[pivotSheetName];
                destination = sheet.Range["A3"];
                pivotTable = pivotCache.CreatePivotTable(destination, connectionName + "_Pivot");
                Assert.NotNull(pivotTable);
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref pivotTable);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref destination);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref sheet);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref sheets);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref pivotCache);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref pivotCaches);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref oledbConnection);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref connection);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref connections);
            }
        });
    }

    private void AssertSchemaAndMemberPaging(
        string connectionName,
        string hierarchyUniqueName,
        string levelUniqueName)
    {
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
