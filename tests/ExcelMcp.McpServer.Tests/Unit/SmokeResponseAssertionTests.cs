using System.Reflection;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;
using Xunit;
using Xunit.Sdk;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
public sealed class SmokeResponseAssertionTests
{
    private static readonly Action<string, string> AssertSmokeSuccess =
        typeof(McpServerSmokeTests)
            .GetMethod("AssertSuccess", BindingFlags.NonPublic | BindingFlags.Static)!
            .CreateDelegate<Action<string, string>>();

    [Theory]
    [InlineData("{}")]
    [InlineData("""{"message":"completed"}""")]
    [InlineData("""{"success":false,"errorMessage":"failed"}""")]
    [InlineData("""{"success":true,"errorMessage":"failed"}""")]
    [InlineData("""{"Success":true,"ErrorMessage":"failed"}""")]
    [InlineData("""{"success":true,"Success":false}""")]
    [InlineData("""{"success":true,"isError":true}""")]
    [InlineData("""{"success":"true"}""")]
    [InlineData("[]")]
    [InlineData("null")]
    [InlineData("not json")]
    public void AssertSuccess_RejectsMissingMalformedOrContradictoryResults(string response)
    {
        Assert.ThrowsAny<XunitException>(() => AssertSmokeSuccess(response, "Synthetic operation"));
    }

    [Theory]
    [InlineData("""{"success":true}""")]
    [InlineData("""{"success":true,"errorMessage":null}""")]
    [InlineData("""{"Success":true,"ErrorMessage":""}""")]
    public void AssertSuccess_AcceptsDocumentedSuccessfulEnvelopes(string response)
    {
        AssertSmokeSuccess(response, "Synthetic operation");
    }

    [Fact]
    public void ReadText_NamedRangeValueWithoutEnvelope_PreservesDocumentedShape()
    {
        const string response = """{"name":"ReportDate","refersTo":"=Data!$C$2","value":45292,"valueType":"Double"}""";
        var result = new CallToolResult
        {
            Content = [new TextContentBlock { Text = response }],
            IsError = false
        };

        Assert.Equal(response, McpResponseAssertions.ReadText(result, expectedError: false));
        Assert.ThrowsAny<XunitException>(() => AssertSmokeSuccess(response, "Envelope operation"));
    }

    [Fact]
    public void ReadText_NonEnvelopeProtocolError_CannotPassAsSuccess()
    {
        var result = new CallToolResult
        {
            Content = [new TextContentBlock { Text = """{"name":"ReportDate","value":45292}""" }],
            IsError = true
        };

        Assert.ThrowsAny<XunitException>(() =>
            McpResponseAssertions.ReadText(result, expectedError: false));
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, true)]
    public void ReadText_FileValidationDiagnostic_PreservesUnsuccessfulAssessment(
        bool exists, bool canOpen, bool isIrmProtected)
    {
        var response = System.Text.Json.JsonSerializer.Serialize(new
        {
            success = false,
            exists,
            isValid = false,
            canOpen,
            isIrmProtected,
            willOpenReadOnly = isIrmProtected,
            requiresVisibleSession = isIrmProtected
        });
        var result = new CallToolResult
        {
            Content = [new TextContentBlock { Text = response }],
            IsError = false
        };

        Assert.Equal(response, McpResponseAssertions.ReadText(result, expectedError: false));
        Assert.ThrowsAny<XunitException>(() => AssertSmokeSuccess(response, "Open workbook"));
    }

    [Theory]
    [InlineData("""{"success":false,"canOpen":false}""")]
    [InlineData("""{"success":false,"exists":false,"isValid":false,"canOpen":"false","isIrmProtected":false,"willOpenReadOnly":false,"requiresVisibleSession":false}""")]
    public void ReadText_IncompleteOrMalformedFileValidation_RejectsDiagnostic(string response)
    {
        var result = new CallToolResult
        {
            Content = [new TextContentBlock { Text = response }],
            IsError = false
        };

        Assert.ThrowsAny<XunitException>(() =>
            McpResponseAssertions.ReadText(result, expectedError: false));
    }

    [Fact]
    public void ReadText_FileValidationDiagnostic_WithProtocolError_IsRejected()
    {
        var result = new CallToolResult
        {
            Content = [new TextContentBlock
            {
                Text = """{"success":false,"exists":false,"isValid":false,"canOpen":false,"isIrmProtected":false,"willOpenReadOnly":false,"requiresVisibleSession":false}"""
            }],
            IsError = true
        };

        Assert.ThrowsAny<XunitException>(() => McpResponseAssertions.ReadText(result));
    }
}
