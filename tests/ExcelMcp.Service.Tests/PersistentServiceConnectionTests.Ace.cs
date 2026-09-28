using System.Diagnostics;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

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
        }
        finally
        {
            File.Delete(sourceWorkbook);
        }
    }

    [Fact]
    public void Refresh_AceOleDbConnection_ReturnsSuccess()
    {
        var (sourceWorkbook, connectionName) = CreateAceConnection();
        var sheetName = UniqueSheetName("ProductsData");
        try
        {
            _connections.LoadTo(_fixture.BatchToken, connectionName, sheetName);
            _fixture.RegisterSheetForCleanup(sheetName);

            var result = _connections.Refresh(
                _fixture.BatchToken,
                connectionName);

            Assert.True(result.Success, result.ErrorMessage);
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
            _connections.LoadTo(_fixture.BatchToken, connectionName, sheetName);
            _fixture.RegisterSheetForCleanup(sheetName);
            AceOleDbTestHelper.UpdateExcelDataSource(sourceWorkbook, sheet =>
            {
                sheet.Range["B2"].Value2 = 49.99;
            });

            var result = _connections.Refresh(
                _fixture.BatchToken,
                connectionName);

            Assert.True(result.Success, result.ErrorMessage);
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
            _connections.LoadTo(_fixture.BatchToken, connectionName, sheetName);
            _fixture.RegisterSheetForCleanup(sheetName);
            _connections.SetProperties(
                _fixture.BatchToken,
                connectionName,
                backgroundQuery: true);
            var before = _connections.GetProperties(
                _fixture.BatchToken,
                connectionName);
            Assert.True(
                before.BackgroundQuery,
                "Precondition: BackgroundQuery should be true before refresh.");

            _connections.Refresh(_fixture.BatchToken, connectionName);

            var after = _connections.GetProperties(
                _fixture.BatchToken,
                connectionName);
            Assert.True(
                after.BackgroundQuery,
                "BackgroundQuery must be restored to true after refresh.");
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
            _connections.LoadTo(_fixture.BatchToken, connectionName, sheetName);
            _fixture.RegisterSheetForCleanup(sheetName);
            _connections.SetProperties(
                _fixture.BatchToken,
                connectionName,
                backgroundQuery: false);
            var before = _connections.GetProperties(
                _fixture.BatchToken,
                connectionName);
            Assert.False(
                before.BackgroundQuery,
                "Precondition: BackgroundQuery should be false.");

            _connections.Refresh(_fixture.BatchToken, connectionName);

            var after = _connections.GetProperties(
                _fixture.BatchToken,
                connectionName);
            Assert.False(
                after.BackgroundQuery,
                "BackgroundQuery must remain false after refresh.");
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

            _connections.LoadTo(_fixture.BatchToken, connectionName, sheetName);
            _fixture.RegisterSheetForCleanup(sheetName);

            stopwatch.Stop();
            Assert.True(
                stopwatch.Elapsed < TimeSpan.FromSeconds(60),
                $"LoadTo took {stopwatch.Elapsed.TotalSeconds:F1}s - suspiciously slow. " +
                "Possible deadlock regression in the connection QueryTable refresh path.");
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
            _connections.LoadTo(_fixture.BatchToken, connectionName, sheetName);
            _fixture.RegisterSheetForCleanup(sheetName);
            var stopwatch = Stopwatch.StartNew();

            _connections.Refresh(_fixture.BatchToken, connectionName);

            stopwatch.Stop();
            Assert.True(
                stopwatch.Elapsed < TimeSpan.FromSeconds(60),
                $"Refresh took {stopwatch.Elapsed.TotalSeconds:F1}s - suspiciously slow. " +
                "Possible deadlock regression in the connection refresh path.");
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
            var status = _connections.GetRefreshStatus(
                _fixture.BatchToken,
                connectionName);
            Assert.True(status.Success);
            Assert.True(status.SupportsRefreshStatus);
            Assert.False(status.IsRefreshing);

            var cancel = _connections.CancelRefresh(
                _fixture.BatchToken,
                connectionName);
            Assert.True(cancel.Success);
            Assert.True(cancel.SupportsCancellation);
            Assert.False(cancel.WasRefreshing);
            Assert.False(cancel.Cancelled);
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
        _connections.Create(
            _fixture.BatchToken,
            connectionName,
            AceOleDbTestHelper.GetExcelConnectionString(sourceWorkbook),
            commandText: AceOleDbTestHelper.GetDefaultCommandText());
        _fixture.RegisterConnectionForCleanup(connectionName);
        return (sourceWorkbook, connectionName);
    }

    private static string UniqueSheetName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..31];
}
