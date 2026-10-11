using Microsoft.ApplicationInsights.DataContracts;

namespace Sbroenne.ExcelMcp.CLI.Telemetry;

/// <summary>Destination for CLI command telemetry.</summary>
internal interface ICliTelemetrySink : IDisposable
{
    void Track(EventTelemetry eventTelemetry, RequestTelemetry requestTelemetry);

    Task FlushAsync();
}
