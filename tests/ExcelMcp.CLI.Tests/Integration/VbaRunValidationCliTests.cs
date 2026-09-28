using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "VBA")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
public sealed class VbaRunValidationCliTests
{
    [Fact]
    public async Task VbaRun_WhitespaceProcedureName_IsRejectedBeforeServiceCall()
    {
        var result = await CliProcessHelper.RunAsync(
            ["vba", "run", "--session", "unused-session", "--procedure-name", "   "],
            timeoutMs: 60000,
            diagnosticLabel: "vba run with whitespace procedure name");

        Assert.Equal(1, result.ExitCode);

        using var json = JsonDocument.Parse(result.Stdout);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains(
            "procedureName is required for run action",
            json.RootElement.GetProperty("error").GetString());
    }
}
