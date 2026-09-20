using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Tests for Connection View/Properties operations
/// </summary>
public partial class ConnectionCommandsTests
{
    [Fact]
    public void View_Credentials_RedactsBothOutputsWithoutChangingConnection()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.CreateTestFile());
        const string connectionName = "CredentialRedaction";
        var connectionString = "ODBC;DSN=UnconfiguredTestSource;UID=synthetic-user;" +
            string.Join("=", "PWD", "synthetic-password") + ";";
        string? original = null;
        batch.Execute((ctx, ct) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.ODBCConnection? odbc = null;
            try
            {
                connections = ctx.Book.Connections;
                connection = connections.Add2(connectionName, "", connectionString, "", Type.Missing, false, false);
                odbc = connection.ODBCConnection;
                original = Convert.ToString(odbc.Connection);
                Assert.Contains("synthetic-password", original);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref odbc);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });

        var result = _commands.View(batch, connectionName);
        Assert.True(result.Success);
        Assert.DoesNotContain("synthetic-user", result.ConnectionString);
        Assert.DoesNotContain("synthetic-password", result.ConnectionString);
        Assert.Contains("(redacted)", result.ConnectionString);
        Assert.DoesNotContain("synthetic-user", result.DefinitionJson);
        Assert.DoesNotContain("synthetic-password", result.DefinitionJson);
        using var definition = JsonDocument.Parse(result.DefinitionJson);
        Assert.Equal(result.ConnectionString, definition.RootElement.GetProperty("ConnectionString").GetString());

        batch.Execute((ctx, ct) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.ODBCConnection? odbc = null;
            try
            {
                connections = ctx.Book.Connections;
                connection = connections.Item(connectionName);
                odbc = connection.ODBCConnection;
                Assert.Equal(original, Convert.ToString(odbc.Connection));
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
    public void View_ExistingConnection_ReturnsDetails()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Use ODBC connection (doesn't need actual DSN for view test)
        var connName = "ViewTestConnection";
        string connectionString = "ODBC;DSN=ViewTestDSN;DBQ=C:\\temp\\viewtest.xlsx";
        ConnectionTestHelper.CreateOdbcConnection(testFile, connName, connectionString);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _commands.View(batch, connName);

        // Assert
        Assert.True(result.Success, $"View failed: {result.ErrorMessage}");
        Assert.Equal(connName, result.ConnectionName);
        Assert.NotNull(result.ConnectionString);
        Assert.NotNull(result.Type);
    }

    [Fact]
    public void View_NonExistentConnection_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act & Assert
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<InvalidOperationException>(() => _commands.View(batch, "NonExistent"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}

