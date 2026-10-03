using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    private sealed record NativeConnection(
        string Name, string Description, Excel.XlConnectionType Type,
        string Source, string Command, bool Background, bool RefreshOnOpen, bool Refreshing,
        bool SavePassword, int RefreshPeriod);

    private NativeConnection ReadNativeConnection(string name) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.ODBCConnection? odbc = null;
            Excel.OLEDBConnection? oledb = null;
            try
            {
                connections = context.Book.Connections;
                connection = connections.Item(name);
                Assert.Equal(name, connection.Name);
                if (connection.Type == Excel.XlConnectionType.xlConnectionTypeODBC)
                {
                    odbc = connection.ODBCConnection;
                    return new NativeConnection(connection.Name, connection.Description, connection.Type,
                        Convert.ToString(odbc.Connection, CultureInfo.InvariantCulture)!,
                        ConnectionCommandText(odbc.CommandText),
                        odbc.BackgroundQuery, odbc.RefreshOnFileOpen, odbc.Refreshing,
                        odbc.SavePassword, odbc.RefreshPeriod);
                }
                Assert.Equal(Excel.XlConnectionType.xlConnectionTypeOLEDB, connection.Type);
                oledb = connection.OLEDBConnection;
                return new NativeConnection(connection.Name, connection.Description, connection.Type,
                    Convert.ToString(oledb.Connection, CultureInfo.InvariantCulture)!,
                    ConnectionCommandText(oledb.CommandText),
                    oledb.BackgroundQuery, oledb.RefreshOnFileOpen, oledb.Refreshing,
                    oledb.SavePassword, oledb.RefreshPeriod);
            }
            finally
            {
                ComUtilities.Release(ref oledb);
                ComUtilities.Release(ref odbc);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });

    private static string ConnectionCommandText(object? value) =>
        value is Array values
            ? string.Join("\n", values.Cast<object>().Select(item =>
                Convert.ToString(item, CultureInfo.InvariantCulture)))
            : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;

    private NativeConnection SeedRetainedConnection()
    {
        var name = UniqueConnectionName("Retained");
        CreateTrackedConnection(name, "ODBC;DSN=RetainedSource");
        RequireSuccess(_connections.SetProperties(
            _fixture.BatchToken, name, description: "Retained configuration",
            backgroundQuery: true, refreshOnFileOpen: false));
        return ReadNativeConnection(name);
    }

    private void AssertRetainedConnection(NativeConnection before)
    {
        Assert.Equal(before, ReadNativeConnection(before.Name));
        var listed = RequireSuccess(_connections.List(_fixture.BatchToken));
        var connection = Assert.Single(listed.Connections, item => item.Name == before.Name);
        Assert.Equal(before.Description, connection.Description);
        var read = RequireSuccess(_connections.View(_fixture.BatchToken, before.Name));
        Assert.Equal(before.Name, read.ConnectionName);
        Assert.Contains("RetainedSource", read.ConnectionString, StringComparison.Ordinal);
    }

    private void AssertPreservedTextData(string sheetName)
    {
        var values = RequireSuccess(_commands.GetValues(
            _fixture.BatchToken, sheetName, "A1:B2")).Values;
        Assert.Equal(2, values.Count);
        Assert.Equal(["Name", "Value"], values[0]);
        Assert.Equal("Preserved", values[1][0]);
        Assert.Equal(1, Convert.ToInt32(values[1][1], CultureInfo.InvariantCulture));
    }
}
