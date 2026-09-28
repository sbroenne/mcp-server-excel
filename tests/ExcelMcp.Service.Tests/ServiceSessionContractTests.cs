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
        Assert.Contains("Unknown session action", response.ErrorMessage, StringComparison.Ordinal);
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
