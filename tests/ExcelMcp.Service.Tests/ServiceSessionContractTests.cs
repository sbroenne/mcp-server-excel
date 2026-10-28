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
    [Theory]
    [InlineData("enable-editing", """{"approved":"true","allowRememberedTrust":true}""", "approved")]
    [InlineData("enable-connections", """{"approved":true,"allowRememberedTrust":1}""", "allowRememberedTrust")]
    [InlineData("enable-macros", """{"approved":true,"allowReopen":null}""", "allowReopen")]
    [InlineData("notices", """{"approved":true}""", "approved")]
    [InlineData("enable-macros", """{"approved":true,"allowRememberedTrust":true}""", "allowRememberedTrust")]
    public async Task SecurityAction_InvalidApprovalIsRejectedBeforeSelectingAWorkbook(
        string action, string args, string invalidParameter)
    {
        using var service = new ExcelMcpService();
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session." + action,
            SessionId = "missing-session",
            Args = args
        });
        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains(invalidParameter, response.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(0, service.SessionCount);
    }

    [Fact]
    public async Task Save_MissingSession_ReturnsSessionNotFound()
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.save",
            SessionId = "missing-session"
        });

        Assert.False(response.Success);
        Assert.Equal("SessionNotFound", response.ErrorCategory);
        Assert.Contains("missing-session", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(0, service.SessionCount);
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

    [Theory]
    [InlineData("session.open", true)]
    [InlineData("session.open", false)]
    [InlineData("session.create", true)]
    [InlineData("session.create", false)]
    public async Task OpenAndCreate_RejectRemovedTracingOption(string command, bool enabled)
    {
        using var service = new ExcelMcpService();

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = command,
            Args = JsonSerializer.Serialize(new { filePath = @"C:\missing\book.xlsx", pqTracing = enabled })
        });

        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Contains("Unknown parameter(s)", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("pqTracing", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(0, service.SessionCount);
    }

    [Fact]
    public async Task Test_RemoteUrl_ReportsUnavailableLocalMetadataWithoutOpeningExcel()
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
        Assert.False(info.RequiresVisibleSession);
        Assert.Equal(".xlsx", info.Extension);
        Assert.Contains("authentication", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("not found", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, service.SessionCount);
    }

    [Fact]
    public async Task Create_RemoteUrl_IsRejectedWithoutStartingExcel()
    {
        using var service = new ExcelMcpService();
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.create",
            Args = """{"filePath":"https://contoso.sharepoint.com/Documents/Test.xlsx","show":true}"""
        });
        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Equal(0, service.SessionCount);
    }
}
