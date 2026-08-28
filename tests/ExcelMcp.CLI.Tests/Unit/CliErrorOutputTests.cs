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
