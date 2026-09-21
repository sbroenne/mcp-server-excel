using Sbroenne.ExcelMcp.ComInterop.ServiceClient;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "ServiceProtocol")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ServiceProtocolTests
{
    [Fact]
    public void Request_WithoutSource_PreservesNull()
    {
        var request = new ServiceRequest { Command = "service.ping" };

        Assert.Null(request.Source);
        Assert.Null(ServiceProtocol.Deserialize<ServiceRequest>(
            """{"command":"service.ping"}""")!.Source);
        Assert.Null(ServiceProtocol.Deserialize<ServiceRequest>(
            """{"command":"service.ping","source":null}""")!.Source);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("cli")]
    [InlineData("mcp")]
    [InlineData("cli-batch")]
    public void Request_RoundTrip_PreservesSourceAndSession(string? source)
    {
        var request = new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = "test-session",
            Args = """{"includeHidden":true}""",
            Source = source
        };

        var actual = ServiceProtocol.Deserialize<ServiceRequest>(ServiceProtocol.Serialize(request))!;

        Assert.Equal(request.Command, actual.Command);
        Assert.Equal(request.SessionId, actual.SessionId);
        Assert.Equal(request.Args, actual.Args);
        Assert.Equal(source, actual.Source);
    }

    [Fact]
    public void Response_RoundTrip_PreservesRichError()
    {
        var response = new ServiceResponse
        {
            Success = false,
            Command = "sheet.list",
            SessionId = "test-session",
            ErrorMessage = "Operation failed.",
            ErrorCategory = "ComError",
            ExceptionType = "COMException",
            HResult = "0x800A03EC",
            InnerError = "Inner failure.",
            Result = """{"success":false}"""
        };

        var json = ServiceProtocol.Serialize(response);
        var actual = ServiceProtocol.Deserialize<ServiceResponse>(json)!;

        Assert.False(actual.Success);
        Assert.Equal(response.Command, actual.Command);
        Assert.Equal(response.SessionId, actual.SessionId);
        Assert.Equal(response.ErrorMessage, actual.ErrorMessage);
        Assert.Equal(response.ErrorCategory, actual.ErrorCategory);
        Assert.Equal(response.ExceptionType, actual.ExceptionType);
        Assert.Equal(response.HResult, actual.HResult);
        Assert.Equal(response.InnerError, actual.InnerError);
        Assert.Equal(response.Result, actual.Result);
        Assert.Contains("\"hresult\":", json);
    }

    [Fact]
    public void Response_RoundTrip_PreservesSuccessWithoutError()
    {
        var response = new ServiceResponse
        {
            Success = true,
            Command = "session.create",
            SessionId = "created-session",
            Result = """{"success":true}"""
        };

        var actual = ServiceProtocol.Deserialize<ServiceResponse>(ServiceProtocol.Serialize(response))!;

        Assert.True(actual.Success);
        Assert.Null(actual.ErrorMessage);
        Assert.Equal(response.Command, actual.Command);
        Assert.Equal(response.SessionId, actual.SessionId);
        Assert.Equal(response.Result, actual.Result);
    }
}
