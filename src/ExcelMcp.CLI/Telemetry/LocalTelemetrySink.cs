using Microsoft.ApplicationInsights.DataContracts;

namespace Sbroenne.ExcelMcp.CLI.Telemetry;

internal sealed class LocalTelemetrySink(TelemetrySettings settings) : ICliTelemetrySink
{
    private readonly LocalTelemetry _telemetry = new(settings, "ExcelMcp.CLI");

    public void Track(EventTelemetry eventTelemetry, RequestTelemetry requestTelemetry)
    {
        _telemetry.Track(eventTelemetry);
        _telemetry.Track(requestTelemetry);
    }

    public Task FlushAsync()
    {
        _telemetry.Flush();
        return Task.CompletedTask;
    }

    public void Dispose() => _telemetry.Dispose();
}
