using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility;

namespace Sbroenne.ExcelMcp.CLI.Telemetry;

/// <summary>
/// Sends CLI command telemetry to Application Insights. Each instance owns its
/// own configuration so nothing is shared through the SDK's process-wide default.
/// </summary>
internal sealed class ApplicationInsightsTelemetrySink : ICliTelemetrySink
{
    // Statsbeat is SDK health data sent to Microsoft, and SDK stats are delivery
    // counters. Neither is CLI usage telemetry, and off Azure each one makes its own
    // blocking 2-second lookup of the Azure VM metadata service.
    private static readonly string[] SdkSelfTelemetryVariables =
    [
        "APPLICATIONINSIGHTS_STATSBEAT_DISABLED",
        "APPLICATIONINSIGHTS_SDKSTATS_DISABLED"
    ];

    private readonly TelemetryConfiguration _configuration;
    private readonly TelemetryClient _client;

    private ApplicationInsightsTelemetrySink(
        TelemetryConfiguration configuration,
        TelemetryClient client)
    {
        _configuration = configuration;
        _client = client;
    }

    internal static ApplicationInsightsTelemetrySink Create(string connectionString)
    {
        DisableSdkSelfTelemetry(Environment.GetEnvironmentVariable, Environment.SetEnvironmentVariable);
        var configuration = CreateConfiguration(connectionString);
        try
        {
            return new ApplicationInsightsTelemetrySink(configuration, new TelemetryClient(configuration));
        }
        catch (Exception)
        {
            configuration.Dispose();
            throw;
        }
    }

    // Live Metrics is never used for a CLI, and it kept a connection to the
    // Live Metrics service open for the life of the process.
    internal static TelemetryConfiguration CreateConfiguration(string connectionString) =>
        new() { ConnectionString = connectionString, EnableLiveMetrics = false };

    /// <summary>
    /// Turns off the SDK's statsbeat and SDK stats features. The SDK reads these
    /// variables when it builds its configuration, so they must be set first.
    /// A value the user already set is respected.
    /// </summary>
    internal static void DisableSdkSelfTelemetry(
        Func<string, string?> getVariable,
        Action<string, string> setVariable)
    {
        foreach (var variable in SdkSelfTelemetryVariables)
        {
            if (string.IsNullOrEmpty(getVariable(variable)))
            {
                setVariable(variable, "true");
            }
        }
    }

    public void Track(EventTelemetry eventTelemetry, RequestTelemetry requestTelemetry)
    {
        _client.TrackEvent(eventTelemetry);
        _client.TrackRequest(requestTelemetry);
    }

    public Task FlushAsync() => _client.FlushAsync(CancellationToken.None);

    public void Dispose() => _configuration.Dispose();
}
