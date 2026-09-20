using Sbroenne.ExcelMcp.CLI.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
public sealed class CliTelemetryTests
{
    [Fact]
    public void CreateCommandInvocationTelemetry_IdentifiesCliEntryPoint()
    {
        var (eventTelemetry, requestTelemetry) =
            CliTelemetry.CreateCommandInvocationTelemetry(
                "range.get-values",
                25,
                succeeded: true,
                errorCategory: null);

        Assert.Equal("range/get-values", eventTelemetry.Name);
        Assert.Equal("cli", eventTelemetry.Properties["EntryPoint"]);
        Assert.Equal("range", eventTelemetry.Properties["Tool"]);
        Assert.Equal("get-values", eventTelemetry.Properties["Action"]);
        Assert.Equal("cli", requestTelemetry.Properties["EntryPoint"]);
        Assert.True(requestTelemetry.Success);
    }

    [Fact]
    public void CreateCommandInvocationTelemetry_DoesNotIncludeFailureDetails()
    {
        var (eventTelemetry, requestTelemetry) =
            CliTelemetry.CreateCommandInvocationTelemetry(
                "session.open",
                25,
                succeeded: false,
                errorCategory: "Permissions");

        Assert.Equal("external-dependency", eventTelemetry.Properties["FailureClass"]);
        Assert.Equal("external-dependency", requestTelemetry.Properties["FailureClass"]);
        Assert.DoesNotContain(
            eventTelemetry.Properties.Keys,
            key => key.Contains("error", StringComparison.OrdinalIgnoreCase));
    }
}
