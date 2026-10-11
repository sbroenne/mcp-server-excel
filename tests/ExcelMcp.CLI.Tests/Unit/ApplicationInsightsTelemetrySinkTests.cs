using Sbroenne.ExcelMcp.CLI.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ApplicationInsightsTelemetrySinkTests
{
    private const string StatsbeatVariable = "APPLICATIONINSIGHTS_STATSBEAT_DISABLED";
    private const string SdkStatsVariable = "APPLICATIONINSIGHTS_SDKSTATS_DISABLED";

    // The ingestion host cannot resolve, so nothing could be sent even by mistake.
    private const string TestConnectionString =
        "InstrumentationKey=00000000-0000-0000-0000-000000000000;IngestionEndpoint=https://localhost.invalid/";

    [Fact]
    public void CreateConfiguration_SetsConnectionStringAndDisablesLiveMetrics()
    {
        // Only the configuration is built here. The SDK starts its network work when a
        // TelemetryClient is created, so no client is constructed in tests.
        using var configuration = ApplicationInsightsTelemetrySink.CreateConfiguration(TestConnectionString);

        Assert.Equal(TestConnectionString, configuration.ConnectionString);
        Assert.False(configuration.EnableLiveMetrics);
    }

    [Fact]
    public void CreateConfiguration_ReturnsIsolatedInstances()
    {
        using var first = ApplicationInsightsTelemetrySink.CreateConfiguration(TestConnectionString);
        using var second = ApplicationInsightsTelemetrySink.CreateConfiguration(TestConnectionString);

        Assert.NotSame(first, second);
    }

    [Fact]
    public void DisableSdkSelfTelemetry_VariablesUnset_SetsBothSwitches()
    {
        var variables = new Dictionary<string, string>();

        ApplicationInsightsTelemetrySink.DisableSdkSelfTelemetry(
            name => variables.GetValueOrDefault(name),
            (name, value) => variables[name] = value);

        Assert.Equal("true", variables[StatsbeatVariable]);
        Assert.Equal("true", variables[SdkStatsVariable]);
    }

    [Fact]
    public void DisableSdkSelfTelemetry_VariablesEmpty_SetsBothSwitches()
    {
        var variables = new Dictionary<string, string>
        {
            [StatsbeatVariable] = "",
            [SdkStatsVariable] = ""
        };

        ApplicationInsightsTelemetrySink.DisableSdkSelfTelemetry(
            name => variables.GetValueOrDefault(name),
            (name, value) => variables[name] = value);

        Assert.Equal("true", variables[StatsbeatVariable]);
        Assert.Equal("true", variables[SdkStatsVariable]);
    }

    [Theory]
    [InlineData(StatsbeatVariable)]
    [InlineData(SdkStatsVariable)]
    public void DisableSdkSelfTelemetry_ExplicitValue_IsLeftUntouched(string explicitVariable)
    {
        var otherVariable = explicitVariable == StatsbeatVariable ? SdkStatsVariable : StatsbeatVariable;
        var variables = new Dictionary<string, string>
        {
            [explicitVariable] = "false"
        };
        var written = new List<string>();

        ApplicationInsightsTelemetrySink.DisableSdkSelfTelemetry(
            name => variables.GetValueOrDefault(name),
            (name, value) =>
            {
                written.Add(name);
                variables[name] = value;
            });

        Assert.Equal("false", variables[explicitVariable]);
        Assert.Equal("true", variables[otherVariable]);
        Assert.Equal([otherVariable], written);
    }
}
