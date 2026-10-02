using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PageLayout")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PageLayoutContractCliTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TypedPageOptions_MapWithoutRequiringOrientation(bool breaks)
    {
        ServiceRequest? captured = null;
        var action = breaks ? "set-page-breaks" : "set-page-setup";
        var payload = breaks ? """{"rows":[10],"columns":[]}""" : """{"printArea":"","leftMargin":36,"zoomPercent":90}""";
        var result = await InProcessCliHelper.RunAsync(
        [
            "-q", "worksheetstyle", action, "--session", "session-1", "--sheet-name", "Report",
            breaks ? "--page-break-options" : "--page-setup-options", payload
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"sheet.{action}", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        var options = args.RootElement.GetProperty(breaks ? "pageBreakOptions" : "pageSetupOptions");
        if (breaks)
        {
            Assert.Equal(10, options.GetProperty("rows")[0].GetInt32());
            Assert.Equal(0, options.GetProperty("columns").GetArrayLength());
        }
        else
        {
            Assert.Equal("", options.GetProperty("printArea").GetString());
            Assert.Equal(36d, options.GetProperty("leftMargin").GetDouble());
            Assert.Equal(90, options.GetProperty("zoomPercent").GetInt32());
        }
    }
}
