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
        Assert.Equal("session-1", captured.SessionId);
        using var document = JsonDocument.Parse(captured.Args!);
        var options = document.RootElement.GetProperty(point ? "pointOptions" : "errorBarOptions");
        using var expected = JsonDocument.Parse(point
            ? """{"fillColor":"#FF0000","markerStyle":"Diamond","markerSize":14}"""
            : """{"enabled":true,"direction":"X","kind":"Fixed","amount":2,"include":"Plus","endStyle":"NoCap"}""");
        Assert.True(JsonElement.DeepEquals(expected.RootElement, options), options.GetRawText());
        Assert.Equal("Sales", document.RootElement.GetProperty("chartName").GetString());
        Assert.Equal(2, document.RootElement.GetProperty("seriesIndex").GetInt32());
        if (point)
            Assert.Equal(3, document.RootElement.GetProperty("pointIndex").GetInt32());
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
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Sales", args.RootElement.GetProperty("chartName").GetString());
        Assert.Equal("sales.jpg", args.RootElement.GetProperty("targetPath").GetString());
        Assert.Equal("Jpeg", args.RootElement.GetProperty("imageFormat").GetString());
        Assert.True(args.RootElement.GetProperty("overwrite").GetBoolean());
    }

    [Theory]
    [InlineData(null, "IOException", "Excel failed to export a nonempty Png image.")]
    [InlineData("Cancelled", "OperationCanceledException", "Chart export was cancelled.")]
    public async Task ImageExport_ForwardsFailuresWithoutSuccessfulResult(string? category, string exception, string message)
    {
        var result = await InProcessCliHelper.RunAsync(
            ["chart", "export-image", "--session", "session-1", "--chart-name", "Sales", "--target-path", "sales.png"], request =>
            {
                Assert.Equal("chart.export-image", request.Command);
                return new ServiceResponse
                {
                    Success = false,
                    Command = request.Command,
                    SessionId = request.SessionId,
                    ErrorCategory = category,
                    ExceptionType = exception,
                    ErrorMessage = message
                };
            });
        Assert.Equal(1, result.ExitCode);
        using var json = JsonDocument.Parse(result.Stdout);
        var root = json.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        if (category is null)
            Assert.False(root.TryGetProperty("errorCategory", out _));
        else
            Assert.Equal(category, root.GetProperty("errorCategory").GetString());
        Assert.Equal(exception, root.GetProperty("exceptionType").GetString());
        Assert.Equal(message, root.GetProperty("errorMessage").GetString());
        Assert.False(root.TryGetProperty("filePath", out _));
    }
}
