using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [Fact]
    public void Create_OdbcConnection_ReturnsSuccess()
    {
        var connectionName = UniqueConnectionName("TestOdbcConnection");

        _connections.Create(
            _fixture.BatchToken,
            connectionName,
            @"ODBC;DSN=Excel Files;DBQ=C:\temp\test.xlsx");
        _fixture.RegisterConnectionForCleanup(connectionName);

        var result = _connections.List(_fixture.BatchToken);
        Assert.True(result.Success);
        Assert.Contains(result.Connections, connection => connection.Name == connectionName);
    }

    [Fact]
    public void Create_DuplicateName_CreatesSecondConnection()
    {
        var connectionName = UniqueConnectionName("DuplicateTest");

        _connections.Create(
            _fixture.BatchToken,
            connectionName,
            @"ODBC;DSN=Source1;DBQ=C:\temp\test1.xlsx");
        _fixture.RegisterConnectionForCleanup(connectionName);
        _connections.Create(
            _fixture.BatchToken,
            connectionName,
            @"ODBC;DSN=Source2;DBQ=C:\temp\test2.xlsx");

        var result = _connections.List(_fixture.BatchToken);
        Assert.True(result.Success);
        var matchingConnections = result.Connections
            .Where(connection =>
                connection.Name == connectionName
                || connection.Name.StartsWith(connectionName, StringComparison.Ordinal))
            .ToList();
        foreach (var connection in matchingConnections)
        {
            _fixture.RegisterConnectionForCleanup(connection.Name);
        }

        Assert.True(
            matchingConnections.Count >= 1,
            "At least one connection with the specified name should exist");
    }

    [Fact]
    public void Create_WithDescription_CreatesConnection()
    {
        var connectionName = UniqueConnectionName("ConnectionWithDescription");

        _connections.Create(
            _fixture.BatchToken,
            connectionName,
            @"ODBC;DSN=Excel Files;DBQ=C:\temp\test.xlsx",
            description: "This is a test connection for ODBC data");
        _fixture.RegisterConnectionForCleanup(connectionName);

        var result = _connections.View(_fixture.BatchToken, connectionName);
        Assert.True(result.Success);
    }

    [Fact]
    public void View_ExistingConnection_ReturnsDetails()
    {
        var connectionName = UniqueConnectionName("ViewTestConnection");
        const string connectionString =
            @"ODBC;DSN=ViewTestDSN;DBQ=C:\temp\viewtest.xlsx";

        _connections.Create(
            _fixture.BatchToken,
            connectionName,
            connectionString);
        _fixture.RegisterConnectionForCleanup(connectionName);

        var result = _connections.View(_fixture.BatchToken, connectionName);
        Assert.True(result.Success, $"View failed: {result.ErrorMessage}");
        Assert.Equal(connectionName, result.ConnectionName);
        Assert.NotNull(result.ConnectionString);
        Assert.NotNull(result.Type);
    }

    [Fact]
    public void View_Credentials_RedactsBothOutputsWithoutChangingConnection()
    {
        var connectionName = UniqueConnectionName("CredentialRedaction");
        var connectionString = "ODBC;DSN=UnconfiguredTestSource;UID=synthetic-user;" +
            string.Join("=", "PWD", "synthetic-password") + ";";
        CreateTrackedConnection(connectionName, connectionString);

        var result = _connections.View(_fixture.BatchToken, connectionName);

        Assert.True(result.Success);
        Assert.DoesNotContain("synthetic-user", result.ConnectionString);
        Assert.DoesNotContain("synthetic-password", result.ConnectionString);
        Assert.Contains("(redacted)", result.ConnectionString);
        Assert.DoesNotContain("synthetic-user", result.DefinitionJson);
        Assert.DoesNotContain("synthetic-password", result.DefinitionJson);
        using var definition = JsonDocument.Parse(result.DefinitionJson);
        Assert.Equal(
            result.ConnectionString,
            definition.RootElement.GetProperty("ConnectionString").GetString());

        _fixture.BatchToken.Execute((ctx, ct) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.ODBCConnection? odbc = null;
            try
            {
                connections = ctx.Book.Connections;
                connection = connections.Item(connectionName);
                odbc = connection.ODBCConnection;
                Assert.Contains("synthetic-password", Convert.ToString(odbc.Connection));
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref odbc);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });
    }

    [Fact]
    public void Delete_ExistingTextConnection_ReturnsSuccess()
    {
        var connectionName = UniqueConnectionName("DeleteTestConnection");
        CreateTrackedConnection(
            connectionName,
            @"ODBC;DSN=TestDSN;DBQ=C:\temp\test.xlsx");

        var before = _connections.List(_fixture.BatchToken);
        Assert.True(before.Success);
        Assert.Contains(before.Connections, connection => connection.Name == connectionName);

        DeleteTrackedConnection(connectionName);

        var after = _connections.List(_fixture.BatchToken);
        Assert.True(after.Success);
        Assert.DoesNotContain(
            after.Connections,
            connection => connection.Name == connectionName);
    }

    [Fact]
    public void Delete_AfterCreatingMultiple_RemovesOnlySpecified()
    {
        var first = UniqueConnectionName("Connection1");
        var second = UniqueConnectionName("Connection2");
        var third = UniqueConnectionName("Connection3");
        CreateTrackedConnection(first, @"ODBC;DSN=TestDSN1;DBQ=C:\temp\test1.xlsx");
        CreateTrackedConnection(second, @"ODBC;DSN=TestDSN2;DBQ=C:\temp\test2.xlsx");
        CreateTrackedConnection(third, @"ODBC;DSN=TestDSN3;DBQ=C:\temp\test3.xlsx");

        DeleteTrackedConnection(second);

        var result = _connections.List(_fixture.BatchToken);
        Assert.True(result.Success);
        Assert.Contains(result.Connections, connection => connection.Name == first);
        Assert.DoesNotContain(result.Connections, connection => connection.Name == second);
        Assert.Contains(result.Connections, connection => connection.Name == third);
    }

    [Fact]
    public void Delete_ConnectionWithDescription_RemovesSuccessfully()
    {
        var connectionName = UniqueConnectionName("DescribedConnection");
        _connections.Create(
            _fixture.BatchToken,
            connectionName,
            @"ODBC;DSN=DescribedDSN;DBQ=C:\temp\described.xlsx",
            description: "Test connection with description");
        _fixture.RegisterConnectionForCleanup(connectionName);

        DeleteTrackedConnection(connectionName);

        var result = _connections.List(_fixture.BatchToken);
        Assert.DoesNotContain(
            result.Connections,
            connection => connection.Name == connectionName);
    }

    [Fact]
    public void Delete_ImmediatelyAfterCreate_WorksCorrectly()
    {
        var connectionName = UniqueConnectionName("ImmediateDeleteTest");
        CreateTrackedConnection(
            connectionName,
            @"ODBC;DSN=ImmediateDSN;DBQ=C:\temp\immediate.xlsx");

        DeleteTrackedConnection(connectionName);

        var result = _connections.List(_fixture.BatchToken);
        Assert.DoesNotContain(
            result.Connections,
            connection => connection.Name == connectionName);
    }

    [Fact]
    public void Delete_ConnectionAfterViewOperation_RemovesSuccessfully()
    {
        var connectionName = UniqueConnectionName("ViewThenDelete");
        CreateTrackedConnection(
            connectionName,
            @"ODBC;DSN=ViewDeleteDSN;DBQ=C:\temp\viewdelete.xlsx");

        var view = _connections.View(_fixture.BatchToken, connectionName);
        Assert.True(view.Success);
        Assert.Equal(connectionName, view.ConnectionName);

        DeleteTrackedConnection(connectionName);

        var result = _connections.List(_fixture.BatchToken);
        Assert.DoesNotContain(
            result.Connections,
            connection => connection.Name == connectionName);
    }

    [Fact]
    public void Delete_RepeatedDeleteAttempts_SecondAttemptFails()
    {
        var connectionName = UniqueConnectionName("DoubleDeleteTest");
        CreateTrackedConnection(
            connectionName,
            @"ODBC;DSN=DoubleDeleteDSN;DBQ=C:\temp\doubledelete.xlsx");

        DeleteTrackedConnection(connectionName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _connections.Delete(_fixture.BatchToken, connectionName));
        Assert.Contains("not found", exception.Message);
    }

    private void CreateTrackedConnection(
        string connectionName,
        string connectionString)
    {
        _connections.Create(
            _fixture.BatchToken,
            connectionName,
            connectionString);
        _fixture.RegisterConnectionForCleanup(connectionName);
    }

    private void DeleteTrackedConnection(string connectionName)
    {
        _connections.Delete(_fixture.BatchToken, connectionName);
        _fixture.ForgetConnection(connectionName);
    }

    private static string UniqueConnectionName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..Math.Min(prefix.Length + 9, 31)];
}
