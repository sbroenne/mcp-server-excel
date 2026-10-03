using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Timelines")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class TimelineContractCliTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NativeControlOptions_KeepCalendarDatesAndCamelCase(bool selection)
    {
        ServiceRequest? captured = null;
        var action = selection ? "set-timeline-selection" : "update-slicer";
        var payload = selection
            ? """{"startDate":"2024-02-01","endDate":"2024-02-29"}"""
            : """{"width":350,"granularity":"Days","showHeader":false}""";
        var result = await InProcessCliHelper.RunAsync(
        [
            "-q", "slicer", action, "--session", "session-1", "--slicer-name", "Dates",
            selection ? "--timeline-selection" : "--slicer-options", payload
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true,"slicer":{"name":"Dates"}}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"slicer.{action}", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        var options = args.RootElement.GetProperty(selection ? "timelineSelection" : "slicerOptions");
        if (selection)
        {
            Assert.Equal(new DateTime(2024, 2, 1), options.GetProperty("startDate").GetDateTime());
            Assert.Equal(new DateTime(2024, 2, 29), options.GetProperty("endDate").GetDateTime());
        }
        else
        {
            Assert.Equal("Days", options.GetProperty("granularity").GetString());
            Assert.False(options.GetProperty("showHeader").GetBoolean());
        }
    }
}
