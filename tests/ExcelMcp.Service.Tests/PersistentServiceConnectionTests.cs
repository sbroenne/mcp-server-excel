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
        var exception = Assert.Throws<NotSupportedException>(() =>
            _connections.Create(
                _fixture.BatchToken,
                "TestTextConnection",
                @"TEXT;C:\temp\test_data.csv"));

        Assert.Contains(
            "TEXT and WEB connections are no longer supported",
            exception.Message);
        Assert.Contains("powerquery", exception.Message);
    }

    [Fact]
    public void Create_WebConnection_ThrowsNotSupportedException()
    {
        var exception = Assert.Throws<NotSupportedException>(() =>
            _connections.Create(
                _fixture.BatchToken,
                "TestWebConnection",
                "URL;https://example.com/data.xml"));

        Assert.Contains(
            "TEXT and WEB connections are no longer supported",
            exception.Message);
        Assert.Contains("powerquery", exception.Message);
    }

    [Fact]
    public void List_EmptyWorkbook_ReturnsSuccessWithEmptyList()
    {
        var result = _connections.List(_fixture.BatchToken);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Connections);
        Assert.Empty(result.Connections);
        Assert.Equal(_fixture.WorkbookPath, result.FilePath);
    }

    [Fact]
    public void View_NonExistentConnection_ThrowsException()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _connections.View(_fixture.BatchToken, "NonExistent"));

        Assert.Contains(
            "not found",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Refresh_ConnectionNotFound_ThrowsException()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _connections.Refresh(
                _fixture.BatchToken,
                "NonExistentConnection"));

        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void Delete_NonExistentConnection_ThrowsException()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _connections.Delete(
                _fixture.BatchToken,
                "NonExistentConnection"));

        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void Delete_EmptyConnectionName_ThrowsHelpfulPublicError()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _connections.Delete(_fixture.BatchToken, string.Empty));

        Assert.Contains(
            "connectionName is required",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }
}
