using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    // Excel keeps an OLAP session open while a PivotTable uses the connection.
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
                Assert.True(pivotCache.OLAP);
                Assert.Equal(connectionName + "_Pivot", pivotTable.Name);
            }
            finally
            {
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivotCaches);
                ComUtilities.Release(ref oledbConnection);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });
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
