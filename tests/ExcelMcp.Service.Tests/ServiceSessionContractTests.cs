using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ServiceSessionContractTests
{
    [Fact]
    public async Task Save_IsRejectedAsUnknownSessionAction()
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = "missing-session"
        });

        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Equal(
            "Unknown action 'save' for command group 'session'. Valid actions: create, open, close, list, test.",
            response.ErrorMessage);
    }

    [Fact]
    public async Task UnknownServiceAction_ListsValidActions()
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest { Command = "service.restart" });

        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Equal(
            "Unknown action 'restart' for command group 'service'. Valid actions: ping, shutdown, status.",
            response.ErrorMessage);
    }

    [Theory]
    [InlineData("session.open")]
    [InlineData("session.create")]
    public async Task OpenAndCreate_RelativePath_ReturnSharedValidationError(string command)
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            Args = """{"filePath":"relative\\book.txt"}"""
        });

        Assert.False(response.Success);
        Assert.Contains(
            "absolute Windows path",
            response.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Test_SharePointUrl_ReportsInteractiveValidationWithoutLocalFileChecks()
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.test",
            Args = """{"filePath":"https://contoso.sharepoint.com/sites/Test/Shared%20Documents/Test.xlsx?web=1"}"""
        });

        Assert.True(response.Success, response.ErrorMessage);
        var info = JsonSerializer.Deserialize<FileValidationInfo>(response.Result!, ServiceProtocol.JsonOptions);
        Assert.NotNull(info);
        Assert.Equal("https://contoso.sharepoint.com/sites/Test/Shared%20Documents/Test.xlsx", info.FilePath);
        Assert.False(info.CanOpen);
        Assert.False(info.Exists);
        Assert.True(info.RequiresVisibleSession);
        Assert.Equal(".xlsx", info.Extension);
        Assert.Contains("authentication", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("not found", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, service.SessionCount);
    }

    [Fact]
    public async Task Open_SharePointUrlWithoutShow_RejectsBeforeStartingExcel()
    {
        using var service = new ExcelMcpService();
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = """{"filePath":"https://contoso.sharepoint.com/Documents/Test.xlsx"}"""
        });
        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains("show=true", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(0, service.SessionCount);
    }

    [Theory]
    [InlineData("session.open", "https://example.com/Test.xlsx")]
    [InlineData("session.open", "https://contoso.sharepoint.com/_layouts/15/Doc.aspx")]
    [InlineData("session.create", "https://contoso.sharepoint.com/Documents/Test.xlsx")]
    public async Task UnsupportedUrl_IsRejectedWithoutStartingExcel(string command, string url)
    {
        using var service = new ExcelMcpService();
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            Args = JsonSerializer.Serialize(new { filePath = url, show = true }, ServiceProtocol.JsonOptions)
        });
        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Equal(0, service.SessionCount);
    }
}
