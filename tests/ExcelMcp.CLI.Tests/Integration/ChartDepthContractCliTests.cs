using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ChartDepth")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ChartDepthContractCliTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TypedPayloads_KeepNativeEnumsAndCamelCase(bool point)
    {
        ServiceRequest? captured = null;
        var action = point ? "set-point-format" : "set-error-bars";
        var args = new List<string> { "-q", "chartconfig", action, "--session", "session-1", "--chart-name", "Sales", "--series-index", "2" };
        if (point) args.AddRange(["--point-index", "3"]);
        args.AddRange(point
            ? ["--point-options", """{"fillColor":"#FF0000","markerStyle":"Diamond","markerSize":14}"""]
            : ["--error-bar-options", """{"direction":"X","kind":"Fixed","amount":2,"include":"Plus","endStyle":"NoCap"}"""]);
        var result = await InProcessCliHelper.RunAsync([.. args], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"chartconfig.{action}", captured.Command);
        using var document = JsonDocument.Parse(captured.Args!);
        var options = document.RootElement.GetProperty(point ? "pointOptions" : "errorBarOptions");
        Assert.Equal(point ? "Diamond" : "X", options.GetProperty(point ? "markerStyle" : "direction").GetString());
        Assert.Equal(2, document.RootElement.GetProperty("seriesIndex").GetInt32());
    }

    [Fact]
    public async Task ImageExport_MapsFormatAndExplicitOverwrite()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "chart", "export-image", "--session", "session-1", "--chart-name", "Sales", "--target-path", "sales.jpg", "--image-format", "Jpeg", "--overwrite", "true"], request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("chart.export-image", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Jpeg", args.RootElement.GetProperty("imageFormat").GetString());
        Assert.True(args.RootElement.GetProperty("overwrite").GetBoolean());
    }
}
