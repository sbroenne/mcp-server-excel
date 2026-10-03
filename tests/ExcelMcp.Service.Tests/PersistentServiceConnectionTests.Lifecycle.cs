using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [Fact]
    public void Create_OdbcConnection_ReturnsSuccess()
    {
        var connectionName = UniqueConnectionName("TestOdbcConnection");

        CreateTrackedConnection(connectionName, @"ODBC;DSN=Excel Files;DBQ=C:\temp\test.xlsx");

        var result = _connections.List(_fixture.BatchToken);
        Assert.True(result.Success);
        Assert.Contains(result.Connections, connection => connection.Name == connectionName);
        RequireSuccess(result);
        var native = ReadNativeConnection(connectionName);
        Assert.Equal(Microsoft.Office.Interop.Excel.XlConnectionType.xlConnectionTypeODBC, native.Type);
        Assert.Contains("DSN=Excel Files", native.Source, StringComparison.Ordinal);
        Assert.Equal("ODBC", RequireSuccess(_connections.View(_fixture.BatchToken, connectionName)).Type);
    }

    [Fact]
    public void Create_DuplicateName_CreatesSecondConnection()
    {
        var connectionName = UniqueConnectionName("DuplicateTest");

        RequireSuccess(_connections.Create(
            _fixture.BatchToken,
            connectionName,
            @"ODBC;DSN=Source1;DBQ=C:\temp\test1.xlsx"));
        _fixture.RegisterConnectionForCleanup(connectionName);
        RequireSuccess(_connections.Create(
            _fixture.BatchToken,
            connectionName,
            @"ODBC;DSN=Source2;DBQ=C:\temp\test2.xlsx"));

        var result = _connections.List(_fixture.BatchToken);
        Assert.True(result.Success);
        RequireSuccess(result);
        var matchingConnections = result.Connections
            .Where(connection =>
                connection.Name == connectionName
                || connection.Name.StartsWith(connectionName, StringComparison.Ordinal))
            .ToList();
        foreach (var connection in matchingConnections)
        {
            _fixture.RegisterConnectionForCleanup(connection.Name);
        }

        Assert.Equal(2, matchingConnections.Count);
        Assert.Equal(2, matchingConnections.Select(connection => connection.Name).Distinct().Count());
        var definitions = matchingConnections.Select(connection =>
            _connections.View(_fixture.BatchToken, connection.Name)).ToArray();
        Assert.All(definitions, definition => Assert.True(definition.Success, definition.ErrorMessage));
        Assert.Contains(definitions, definition => definition.ConnectionString?.Contains("Source1", StringComparison.Ordinal) == true);
        Assert.Contains(definitions, definition => definition.ConnectionString?.Contains("Source2", StringComparison.Ordinal) == true);
        foreach (var definition in definitions)
        {
            RequireSuccess(definition);
            var native = ReadNativeConnection(definition.ConnectionName);
            Assert.Equal(definition.ConnectionString, native.Source);
        }
    }

    [Fact]
    public void Create_WithDescription_CreatesConnection()
    {
        var connectionName = UniqueConnectionName("ConnectionWithDescription");

        RequireSuccess(_connections.Create(
            _fixture.BatchToken,
            connectionName,
            @"ODBC;DSN=Excel Files;DBQ=C:\temp\test.xlsx",
            description: "This is a test connection for ODBC data"));
        _fixture.RegisterConnectionForCleanup(connectionName);

        var result = _connections.View(_fixture.BatchToken, connectionName);
        Assert.True(result.Success);
        var listed = _connections.List(_fixture.BatchToken);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal("This is a test connection for ODBC data",
            Assert.Single(listed.Connections, connection => connection.Name == connectionName).Description);
        RequireSuccess(result);
        RequireSuccess(listed);
        Assert.Equal("This is a test connection for ODBC data", ReadNativeConnection(connectionName).Description);
    }

    [Fact]
    public void View_ExistingConnection_ReturnsDetails()
    {
        var connectionName = UniqueConnectionName("ViewTestConnection");
        const string sensitiveCredential = "boundary-sensitive-user";
        const string connectionString =
            $@"ODBC;DSN=ViewTestDSN;DBQ=C:\temp\viewtest.xlsx;UID={sensitiveCredential}";

        RequireSuccess(_connections.Create(
            _fixture.BatchToken,
            connectionName,
            connectionString));
        _fixture.RegisterConnectionForCleanup(connectionName);

        var result = _connections.View(_fixture.BatchToken, connectionName);
        Assert.True(result.Success, $"View failed: {result.ErrorMessage}");
        Assert.Equal(connectionName, result.ConnectionName);
        Assert.NotNull(result.ConnectionString);
        Assert.DoesNotContain(sensitiveCredential, result.ConnectionString, StringComparison.Ordinal);
        Assert.Contains("(redacted)", result.ConnectionString, StringComparison.Ordinal);
        Assert.NotNull(result.DefinitionJson);
        Assert.DoesNotContain(sensitiveCredential, result.DefinitionJson, StringComparison.Ordinal);
        Assert.Contains("(redacted)", result.DefinitionJson, StringComparison.Ordinal);
        Assert.NotNull(result.Type);
        RequireSuccess(result);
        Assert.Equal("ODBC", result.Type);
        Assert.Contains(sensitiveCredential, ReadNativeConnection(connectionName).Source, StringComparison.Ordinal);
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
        RequireSuccess(before);

        DeleteTrackedConnection(connectionName);

        var after = _connections.List(_fixture.BatchToken);
        Assert.True(after.Success);
        Assert.DoesNotContain(
            after.Connections,
            connection => connection.Name == connectionName);
        RequireSuccess(after);
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
        var firstBefore = ReadNativeConnection(first);
        var thirdBefore = ReadNativeConnection(third);

        DeleteTrackedConnection(second);

        var result = _connections.List(_fixture.BatchToken);
        Assert.True(result.Success);
        Assert.Contains(result.Connections, connection => connection.Name == first);
        Assert.DoesNotContain(result.Connections, connection => connection.Name == second);
        Assert.Contains(result.Connections, connection => connection.Name == third);
        RequireSuccess(result);
        Assert.Equal(2, result.Connections.Count);
        Assert.Equal(firstBefore, ReadNativeConnection(first));
        Assert.Equal(thirdBefore, ReadNativeConnection(third));
    }

    [Fact]
    public void Delete_ConnectionWithDescription_RemovesSuccessfully()
    {
        var connectionName = UniqueConnectionName("DescribedConnection");
        RequireSuccess(_connections.Create(
            _fixture.BatchToken,
            connectionName,
            @"ODBC;DSN=DescribedDSN;DBQ=C:\temp\described.xlsx",
            description: "Test connection with description"));
        _fixture.RegisterConnectionForCleanup(connectionName);

        DeleteTrackedConnection(connectionName);

        var result = _connections.List(_fixture.BatchToken);
        RequireSuccess(result);
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
        RequireSuccess(result);
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
        RequireSuccess(view);

        DeleteTrackedConnection(connectionName);

        var result = _connections.List(_fixture.BatchToken);
        RequireSuccess(result);
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
        var retained = SeedRetainedConnection();

        DeleteTrackedConnection(connectionName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _connections.Delete(_fixture.BatchToken, connectionName));
        Assert.Contains("not found", exception.Message);
        AssertRetainedConnection(retained);
    }

    private void CreateTrackedConnection(
        string connectionName,
        string connectionString)
    {
        RequireSuccess(_connections.Create(
            _fixture.BatchToken,
            connectionName,
            connectionString));
        _fixture.RegisterConnectionForCleanup(connectionName);
        var native = ReadNativeConnection(connectionName);
        Assert.Equal(Microsoft.Office.Interop.Excel.XlConnectionType.xlConnectionTypeODBC, native.Type);
        Assert.Equal(connectionString, native.Source);
    }

    [Fact]
    public void SetProperties_UpdatesNativeConfigurationAndRetainsOmittedSettings()
    {
        var before = SeedRetainedConnection();
        var batch = _fixture.BatchToken;

        RequireSuccess(_connections.SetProperties(batch, before.Name,
            commandText: "SELECT * FROM [Products]", description: "Updated configuration",
            backgroundQuery: false, refreshOnFileOpen: true, savePassword: false, refreshPeriod: 5));

        var expected = before with
        {
            Command = "SELECT * FROM [Products]",
            Description = "Updated configuration",
            Background = false,
            RefreshOnOpen = true,
            SavePassword = false,
            RefreshPeriod = 5
        };
        Assert.Equal(expected, ReadNativeConnection(before.Name));
        var properties = RequireSuccess(_connections.GetProperties(batch, before.Name));
        Assert.False(properties.BackgroundQuery);
        Assert.True(properties.RefreshOnFileOpen);
        Assert.False(properties.SavePassword);
        Assert.Equal(5, properties.RefreshPeriod);

        RequireSuccess(_connections.SetProperties(batch, before.Name, description: "Description only"));
        Assert.Equal(expected with { Description = "Description only" }, ReadNativeConnection(before.Name));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(int.MinValue)]
    public void SetProperties_NegativeRefreshPeriod_PreservesExistingConfiguration(int period)
    {
        var before = SeedRetainedConnection();

        var error = Record.Exception(() => _connections.SetProperties(
            _fixture.BatchToken, before.Name, description: "Must not apply",
            backgroundQuery: false, refreshPeriod: period));

        Assert.NotNull(error);
        AssertRetainedConnection(before);
        Assert.IsType<ArgumentOutOfRangeException>(error);
        Assert.Contains("refreshPeriod", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_ConfiguredConnection_DoesNotRefreshOrChangeConfiguration()
    {
        var before = SeedRetainedConnection();

        var result = RequireSuccess(_connections.Test(_fixture.BatchToken, before.Name));

        Assert.Equal(_fixture.WorkbookPath, result.FilePath);
        AssertRetainedConnection(before);
        Assert.False(ReadNativeConnection(before.Name).Refreshing);
    }

    private void DeleteTrackedConnection(string connectionName)
    {
        RequireSuccess(_connections.Delete(_fixture.BatchToken, connectionName));
        _fixture.ForgetConnection(connectionName);
    }

    private static string UniqueConnectionName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..Math.Min(prefix.Length + 9, 31)];
}
