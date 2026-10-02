using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "FineFormatting")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class FineFormattingContractCliTests
{
    [Fact]
    public async Task Format_ForwardsMultipleTargetsAndNestedNativeSettings()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "rangeformat", "format", "--session", "session-1", "--sheet-name", "Data",
             "--range-addresses", "A1:B2", "--range-addresses", "D1:E2", "--format-options",
             """{"bold":false,"fontThemeColor":5,"indentLevel":2,"borders":[{"position":"DiagonalUp","lineStyle":"dash","themeColor":6}]}"""],
            request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeformat.format", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(2, args.RootElement.GetProperty("rangeAddresses").GetArrayLength());
        var options = args.RootElement.GetProperty("formatOptions");
        Assert.False(options.GetProperty("bold").GetBoolean());
        Assert.Equal(5, options.GetProperty("fontThemeColor").GetInt32());
        Assert.Equal(2, options.GetProperty("indentLevel").GetInt32());
        Assert.Equal("DiagonalUp", options.GetProperty("borders")[0].GetProperty("position").GetString());
    }

    [Theory]
    [InlineData("format-range")]
    [InlineData("format-ranges")]
    public async Task ObsoleteCommands_DoNotDispatch(string action)
    {
        bool dispatched = false;
        var result = await InProcessCliHelper.RunAsync(["-q", "rangeformat", action, "--session", "session-1"],
            _ =>
            {
                dispatched = true;
                return new ServiceResponse { Success = true };
            });
        Assert.NotEqual(0, result.ExitCode);
        Assert.False(dispatched);
    }
}
