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
}
