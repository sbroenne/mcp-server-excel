using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [Theory]
    [InlineData("ODBC")]
    [InlineData("OLEDB")]
    [InlineData("OLAP")]
    [InlineData("PowerQuery")]
    public void ListAndView_UnsetNativeMetadata_DoNotInventValuesOrFail(string kind)
    {
        var name = UniqueConnectionName("NativeMetadata");
        if (kind == "PowerQuery")
        {
            _fixture.Send("powerquery.create",
                new { queryName = name, mCode = "#table({\"Value\"}, {{7}})", targetSheet = name });
            _fixture.RegisterPowerQueryForCleanup(name);
            _fixture.RegisterSheetForCleanup(name);
        }
        else
        {
            var connectionString = kind switch
            {
                "ODBC" => "ODBC;DSN=UnconfiguredNativeMetadata",
                "OLEDB" => "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=UnconfiguredNativeMetadata",
                _ => "OLEDB;Provider=MSOLAP.8;Data Source=UnconfiguredNativeMetadata;Initial Catalog=Fixture"
            };
            RequireSuccess(_connections.Create(_fixture.BatchToken, name, connectionString,
                kind == "OLAP" ? "Model" : "SELECT * FROM Fixture"));
            _fixture.RegisterConnectionForCleanup(name);
        }

        var native = _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.Ranges? ranges = null;
            Excel.ODBCConnection? odbc = null;
            Excel.OLEDBConnection? oledb = null;
            try
            {
                connections = context.Book.Connections;
                if (kind == "PowerQuery")
                {
                    for (int index = 1; index <= connections.Count; index++)
                    {
                        connection = connections.Item(index);
                        if (PowerQueryHelpers.TryGetMashupLocation(connection, out string location) &&
                            string.Equals(location, name, StringComparison.Ordinal))
                            break;
                        ComUtilities.Release(ref connection);
                    }
                }
                else
                    connection = connections.Item(name);
                Assert.NotNull(connection);
                if (connection.Type == Excel.XlConnectionType.xlConnectionTypeODBC)
                    odbc = connection.ODBCConnection;
                else
                    oledb = connection.OLEDBConnection;
                static string Read(Func<object?> read)
                {
                    try
                    {
                        object? value = read();
                        return value is null ? "<null>" : $"{value.GetType().Name}:{Convert.ToString(value, CultureInfo.InvariantCulture)}";
                    }
                    catch (COMException error)
                    {
                        return $"COM:0x{error.HResult:X8}";
                    }
                }
                ranges = connection.Ranges;
                var states = new[]
                {
                    $"Ranges={ranges.Count}",
                    $"BackgroundQuery={Read(() => odbc is not null ? odbc.BackgroundQuery : oledb!.BackgroundQuery)}",
                    $"RefreshOnFileOpen={Read(() => odbc is not null ? odbc.RefreshOnFileOpen : oledb!.RefreshOnFileOpen)}",
                    $"SavePassword={Read(() => odbc is not null ? odbc.SavePassword : oledb!.SavePassword)}",
                    $"RefreshPeriod={Read(() => odbc is not null ? odbc.RefreshPeriod : oledb!.RefreshPeriod)}",
                    $"RefreshDate={Read(() => odbc is not null ? odbc.RefreshDate : oledb!.RefreshDate)}",
                    $"IsConnected={(kind == "ODBC" ? "<not an OLE DB property>" : Read(() => oledb!.IsConnected))}"
                };
                return (connection.Name, Detail: string.Join("; ", states));
            }
            finally
            {
                ComUtilities.Release(ref oledb);
                ComUtilities.Release(ref odbc);
                ComUtilities.Release(ref ranges);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });
        Assert.Contains("RefreshDate=COM:0x800A03EC", native.Detail, StringComparison.Ordinal);
        Assert.Contains(kind == "ODBC" ? "Ranges=0" : "IsConnected=Boolean:False", native.Detail, StringComparison.Ordinal);
        ConnectionListResult? listed = null;
        var listError = Record.Exception(() => listed = _connections.List(_fixture.BatchToken));
        Assert.True(listError is null, $"{kind}: {native.Detail}; {listError}");
        var info = Assert.Single(RequireSuccess(listed!).Connections, item => item.Name == native.Name);
        Assert.Equal(kind == "ODBC" ? "ODBC" : "OLEDB", info.Type);
        Assert.Equal(kind == "PowerQuery", info.IsPowerQuery);
        Assert.Null(info.LastRefresh);
        RequireSuccess(_connections.View(_fixture.BatchToken, native.Name));
    }
}
