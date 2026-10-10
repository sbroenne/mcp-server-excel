using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServiceConnectionTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IConnectionCommands _connections =
        ServiceCommandProxy.Create<IConnectionCommands>(fixture);

    [Fact]
    public void Create_TextConnection_ThrowsNotSupportedException()
    {
        var retained = SeedRetainedConnection();
        var exception = Assert.Throws<NotSupportedException>(() =>
            _connections.Create(
                _fixture.BatchToken,
                "TestTextConnection",
                @"TEXT;C:\temp\test_data.csv"));

        Assert.Contains(
            "TEXT and WEB connections are no longer supported",
            exception.Message);
        Assert.Contains("powerquery", exception.Message);
        AssertRetainedConnection(retained);
    }

    [Fact]
    public void Create_WebConnection_ThrowsNotSupportedException()
    {
        var retained = SeedRetainedConnection();
        var exception = Assert.Throws<NotSupportedException>(() =>
            _connections.Create(
                _fixture.BatchToken,
                "TestWebConnection",
                "URL;https://example.com/data.xml"));

        Assert.Contains(
            "TEXT and WEB connections are no longer supported",
            exception.Message);
        Assert.Contains("powerquery", exception.Message);
        AssertRetainedConnection(retained);
    }

    [Fact]
    public void List_EmptyWorkbook_ReturnsSuccessWithEmptyList()
    {
        var result = _connections.List(_fixture.BatchToken);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Connections);
        Assert.Empty(result.Connections);
        Assert.Equal(_fixture.WorkbookPath, result.FilePath);
        RequireSuccess(result);
    }

    [Fact]
    public void View_NonExistentConnection_ThrowsException()
    {
        var retained = SeedRetainedConnection();
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _connections.View(_fixture.BatchToken, "NonExistent"));

        Assert.Contains(
            "not found",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        AssertRetainedConnection(retained);
    }

    [Fact]
    public void Refresh_ConnectionNotFound_ThrowsException()
    {
        var retained = SeedRetainedConnection();
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _connections.Refresh(
                _fixture.BatchToken,
                "NonExistentConnection"));

        Assert.Contains("not found", exception.Message);
        AssertRetainedConnection(retained);
    }

    [Fact]
    public void Delete_NonExistentConnection_ThrowsException()
    {
        var retained = SeedRetainedConnection();
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _connections.Delete(
                _fixture.BatchToken,
                "NonExistentConnection"));

        Assert.Contains("not found", exception.Message);
        AssertRetainedConnection(retained);
    }

    [Fact]
    public void Delete_EmptyConnectionName_ThrowsHelpfulPublicError()
    {
        var retained = SeedRetainedConnection();
        var exception = Assert.Throws<ArgumentException>(() =>
            _connections.Delete(_fixture.BatchToken, string.Empty));

        Assert.Contains(
            "connectionName is required",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        AssertRetainedConnection(retained);
    }

    [Theory]
    [InlineData("get-properties")]
    [InlineData("set-properties")]
    [InlineData("get-refresh-status")]
    [InlineData("cancel-refresh")]
    [InlineData("load-to")]
    [InlineData("test")]
    [InlineData("discover-olap-schema")]
    [InlineData("search-olap-members")]
    public void MissingConnection_AdditionalActions_PreserveConfigurationAndCells(string action)
    {
        var retained = SeedRetainedConnection();
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheet, "A1:B2",
            [["Name", "Value"], ["Preserved", 1]]));
        Action request = action switch
        {
            "get-properties" => () => _connections.GetProperties(batch, "MissingConnection"),
            "set-properties" => () => _connections.SetProperties(
                batch, "MissingConnection", description: "Must not apply", backgroundQuery: false),
            "get-refresh-status" => () => _connections.GetRefreshStatus(batch, "MissingConnection"),
            "cancel-refresh" => () => _connections.CancelRefresh(batch, "MissingConnection"),
            "load-to" => () => _connections.LoadTo(batch, "MissingConnection", sheet),
            "test" => () => _connections.Test(batch, "MissingConnection"),
            "discover-olap-schema" => () => _connections.DiscoverOlapSchema(batch, "MissingConnection"),
            "search-olap-members" => () => _connections.SearchOlapMembers(
                batch, "MissingConnection", "[Date].[Calendar]", "[Date].[Calendar].[Month]"),
            _ => throw new ArgumentOutOfRangeException(nameof(action))
        };

        var error = Assert.Throws<InvalidOperationException>(request);

        Assert.Contains("not found", error.Message, StringComparison.OrdinalIgnoreCase);
        AssertRetainedConnection(retained);
        AssertPreservedTextData(sheet);
    }
}
