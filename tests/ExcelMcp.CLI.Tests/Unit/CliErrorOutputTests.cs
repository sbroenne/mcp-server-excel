using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "ErrorHandling")]
[Trait("Speed", "Fast")]
[Collection("ConsoleOutput")]
public sealed class CliErrorOutputTests
{
    [Theory]
    [InlineData(null, "Cancelled")]
    [InlineData("InvalidInput", "InvalidInput")]
    public void WriteException_UsesSharedCategoryUnlessExplicitlyOverridden(string? supplied, string expected)
    {
        using var stdout = new StringWriter();
        var originalOut = Console.Out;
        try
        {
            Console.SetOut(stdout);
            Assert.Equal(1, CliErrorOutput.WriteException(
                new InvalidOperationException("Operation context", new OperationCanceledException("Cancelled")),
                supplied));
        }
        finally
        {
            Console.SetOut(originalOut);
        }

        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal(expected, json.RootElement.GetProperty("errorCategory").GetString());
        Assert.Equal("Operation context", json.RootElement.GetProperty("errorMessage").GetString());
    }

    [Theory]
    [InlineData("Prerequisite")]
    [InlineData("DependencyUnavailable")]
    [InlineData("Permissions")]
    [InlineData("NotFound")]
    [InlineData("ComInterop")]
    public void SerializeServiceError_PreservesOperationFailureContext(string category)
    {
        using var stdout = new StringWriter();
        var originalOut = Console.Out;
        try
        {
            Console.SetOut(stdout);
            Assert.Equal(1, CliErrorOutput.WriteServiceError(new ServiceResponse
            {
                Success = false,
                ErrorCategory = category,
                ErrorMessage = "Original operation context",
                ExceptionType = "InvalidOperationException",
                InnerError = "Original cause",
                Command = "datamodel.evaluate",
                SessionId = "test-session",
                HResult = "0x800A03EC"
            }));
        }
        finally
        {
            Console.SetOut(originalOut);
        }

        using var json = JsonDocument.Parse(stdout.ToString());
        var root = json.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal(category, root.GetProperty("errorCategory").GetString());
        Assert.Equal("Original operation context", root.GetProperty("errorMessage").GetString());
        Assert.Equal("Original cause", root.GetProperty("innerError").GetString());
        Assert.Equal("datamodel.evaluate", root.GetProperty("command").GetString());
        Assert.Equal("test-session", root.GetProperty("sessionId").GetString());
        Assert.Equal("0x800A03EC", root.GetProperty("hresult").GetString());
    }

    [Fact]
    public void WriteServiceError_PreservesCategoryAndRecoveryGuidance()
    {
        using var stdout = new StringWriter();
        var originalOut = Console.Out;

        try
        {
            Console.SetOut(stdout);
            var exitCode = CliErrorOutput.WriteServiceError(new ServiceResponse
            {
                Success = false,
                ErrorCategory = "SessionNotFound",
                ErrorMessage = "Session not found. Open the workbook again."
            });

            Assert.Equal(1, exitCode);
        }
        finally
        {
            Console.SetOut(originalOut);
        }

        using var json = JsonDocument.Parse(stdout.ToString());
        Assert.Equal(
            "SessionNotFound",
            json.RootElement.GetProperty("errorCategory").GetString());
        Assert.Contains(
            "Open the workbook again",
            json.RootElement.GetProperty("errorMessage").GetString());
    }
}
