using System.Diagnostics;
using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [Fact]
    public void Create_AceOleDbConnection_ReturnsSuccess()
    {
        var (sourceWorkbook, connectionName) = CreateAceConnection();

        try
        {
            var listResult = _connections.List(_fixture.BatchToken);
            Assert.True(listResult.Success);
            Assert.Contains(
                listResult.Connections,
                connection => connection.Name == connectionName);
            RequireSuccess(listResult);
            Assert.False(ReadNativeConnection(connectionName).Refreshing);
        }
        finally
        {
            File.Delete(sourceWorkbook);
        }
    }

    [Fact]
    public void Refresh_AceOleDbConnectionAfterDataUpdate_ReturnsSuccess()
    {
        var (sourceWorkbook, connectionName) = CreateAceConnection();
        var sheetName = UniqueSheetName("ProductsData");
        try
        {
            RequireSuccess(_connections.LoadTo(_fixture.BatchToken, connectionName, sheetName));
            _fixture.RegisterSheetForCleanup(sheetName);
            StageAceDataUpdate(sourceWorkbook, sheetName);

            var result = _connections.Refresh(
                _fixture.BatchToken,
                connectionName);

            Assert.True(result.Success, result.ErrorMessage);
            RequireSuccess(result);
            Assert.False(ReadNativeConnection(connectionName).Refreshing);
            AssertAceData(sheetName, 49.99);
        }
        finally
        {
            File.Delete(sourceWorkbook);
        }
    }

    [Fact]
    public void Refresh_BackgroundQueryTrue_RestoredAfterRefresh()
    {
        var (sourceWorkbook, connectionName) = CreateAceConnection();
        var sheetName = UniqueSheetName("ProductsData");
        try
        {
            RequireSuccess(_connections.LoadTo(_fixture.BatchToken, connectionName, sheetName));
            _fixture.RegisterSheetForCleanup(sheetName);
            RequireSuccess(_connections.SetProperties(
                _fixture.BatchToken,
                connectionName,
                backgroundQuery: true));
            var before = RequireSuccess(_connections.GetProperties(
                _fixture.BatchToken,
                connectionName));
            Assert.True(
                before.BackgroundQuery,
                "Precondition: BackgroundQuery should be true before refresh.");
            StageAceDataUpdate(sourceWorkbook, sheetName);

            RequireSuccess(_connections.Refresh(_fixture.BatchToken, connectionName));

            var after = RequireSuccess(_connections.GetProperties(
                _fixture.BatchToken,
                connectionName));
            Assert.True(
                after.BackgroundQuery,
                "BackgroundQuery must be restored to true after refresh.");
            AssertAceData(sheetName, 49.99);
            var native = ReadNativeConnection(connectionName);
            Assert.True(native.Background);
            Assert.False(native.Refreshing);
        }
        finally
        {
            File.Delete(sourceWorkbook);
        }
    }

    [Fact]
    public void Refresh_BackgroundQueryFalse_RemainsAfterRefresh()
    {
        var (sourceWorkbook, connectionName) = CreateAceConnection();
        var sheetName = UniqueSheetName("ProductsData");
        try
        {
            RequireSuccess(_connections.LoadTo(_fixture.BatchToken, connectionName, sheetName));
            _fixture.RegisterSheetForCleanup(sheetName);
            RequireSuccess(_connections.SetProperties(
                _fixture.BatchToken,
                connectionName,
                backgroundQuery: false));
            var before = RequireSuccess(_connections.GetProperties(
                _fixture.BatchToken,
                connectionName));
            Assert.False(
                before.BackgroundQuery,
                "Precondition: BackgroundQuery should be false.");
            StageAceDataUpdate(sourceWorkbook, sheetName);

            RequireSuccess(_connections.Refresh(_fixture.BatchToken, connectionName));

            var after = RequireSuccess(_connections.GetProperties(
                _fixture.BatchToken,
                connectionName));
            Assert.False(
                after.BackgroundQuery,
                "BackgroundQuery must remain false after refresh.");
            AssertAceData(sheetName, 49.99);
            var native = ReadNativeConnection(connectionName);
            Assert.False(native.Background);
            Assert.False(native.Refreshing);
        }
        finally
        {
            File.Delete(sourceWorkbook);
        }
    }

    [Fact]
    public void LoadTo_AceOleDbConnection_CompletesWithoutDeadlock()
    {
        var (sourceWorkbook, connectionName) = CreateAceConnection();
        var sheetName = UniqueSheetName("ProductsData");
        try
        {
            var stopwatch = Stopwatch.StartNew();

            RequireSuccess(_connections.LoadTo(_fixture.BatchToken, connectionName, sheetName));
            _fixture.RegisterSheetForCleanup(sheetName);

            stopwatch.Stop();
            Assert.True(
                stopwatch.Elapsed < TimeSpan.FromSeconds(60),
                $"LoadTo took {stopwatch.Elapsed.TotalSeconds:F1}s - suspiciously slow. " +
                "Possible deadlock regression in the connection QueryTable refresh path.");
            AssertAceData(sheetName, 19.99);
            Assert.False(ReadNativeConnection(connectionName).Refreshing);
        }
        finally
        {
            File.Delete(sourceWorkbook);
        }
    }

    [Fact]
    public void Refresh_AceOleDbConnection_CompletesWithoutDeadlock()
    {
        var (sourceWorkbook, connectionName) = CreateAceConnection();
        var sheetName = UniqueSheetName("ProductsData");
        try
        {
            RequireSuccess(_connections.LoadTo(_fixture.BatchToken, connectionName, sheetName));
            _fixture.RegisterSheetForCleanup(sheetName);
            StageAceDataUpdate(sourceWorkbook, sheetName);
            var stopwatch = Stopwatch.StartNew();

            RequireSuccess(_connections.Refresh(_fixture.BatchToken, connectionName));

            stopwatch.Stop();
            Assert.True(
                stopwatch.Elapsed < TimeSpan.FromSeconds(60),
                $"Refresh took {stopwatch.Elapsed.TotalSeconds:F1}s - suspiciously slow. " +
                "Possible deadlock regression in the connection refresh path.");
            AssertAceData(sheetName, 49.99);
            Assert.False(ReadNativeConnection(connectionName).Refreshing);
        }
        finally
        {
            File.Delete(sourceWorkbook);
        }
    }

    [Fact]
    public void RefreshControl_IdleOleDbConnection_ReportsStatusAndDoesNotCancel()
    {
        var (sourceWorkbook, connectionName) = CreateAceConnection();
        try
        {
            var before = ReadNativeConnection(connectionName);
            var status = RequireSuccess(_connections.GetRefreshStatus(
                _fixture.BatchToken,
                connectionName));
            Assert.True(status.Success);
            Assert.True(status.SupportsRefreshStatus);
            Assert.False(status.IsRefreshing);

            var cancel = RequireSuccess(_connections.CancelRefresh(
                _fixture.BatchToken,
                connectionName));
            Assert.True(cancel.Success);
            Assert.True(cancel.SupportsCancellation);
            Assert.False(cancel.WasRefreshing);
            Assert.False(cancel.Cancelled);
            Assert.Equal(before, ReadNativeConnection(connectionName));
        }
        finally
        {
            File.Delete(sourceWorkbook);
        }
    }

    private (string SourceWorkbook, string ConnectionName) CreateAceConnection()
    {
        var sourceWorkbook = Path.Combine(
            Path.GetDirectoryName(_fixture.WorkbookPath)!,
            $"AceSource_{Guid.NewGuid():N}.xlsx");
        AceOleDbTestHelper.CreateExcelDataSource(sourceWorkbook);
        var connectionName = UniqueConnectionName("AceOleDb");
        RequireSuccess(_connections.Create(
            _fixture.BatchToken,
            connectionName,
            AceOleDbTestHelper.GetExcelConnectionString(sourceWorkbook),
            commandText: AceOleDbTestHelper.GetDefaultCommandText()));
        _fixture.RegisterConnectionForCleanup(connectionName);
        var native = ReadNativeConnection(connectionName);
        Assert.Equal(Excel.XlConnectionType.xlConnectionTypeOLEDB, native.Type);
        Assert.Contains(sourceWorkbook, native.Source, StringComparison.Ordinal);
        Assert.Equal(AceOleDbTestHelper.GetDefaultCommandText(), native.Command);
        return (sourceWorkbook, connectionName);
    }

    private static string UniqueSheetName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..31];

    private void StageAceDataUpdate(string sourceWorkbook, string sheetName)
    {
        AssertAceData(sheetName, 19.99);
        AceOleDbTestHelper.UpdateExcelDataSource(sourceWorkbook, sheet =>
        {
            Excel.Range? cell = null;
            try
            {
                cell = ((Excel.Worksheet)sheet).Range["B2"];
                cell.Value2 = 49.99;
            }
            finally { ComUtilities.Release(ref cell); }
        });
        AssertAceData(sheetName, 19.99);
    }

    private void AssertAceData(string sheetName, double expectedPrice)
    {
        var result = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:B3");
        Assert.True(result.Success, result.ErrorMessage);
        RequireSuccess(result);
        Assert.Equal(3, result.Values.Count);
        Assert.All(result.Values, row => Assert.Equal(2, row.Count));
        Assert.Equal(["Product", "Price"], result.Values[0]);
        Assert.Equal("Widget", result.Values[1][0]);
        Assert.Equal(expectedPrice, Convert.ToDouble(result.Values[1][1], CultureInfo.InvariantCulture));
        Assert.Equal("Gadget", result.Values[2][0]);
        Assert.Equal(29.99, Convert.ToDouble(result.Values[2][1], CultureInfo.InvariantCulture));
    }
}
