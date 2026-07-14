using System.Diagnostics.Metrics;
using OpenTelemetry;
using OpenTelemetry.Metrics;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

/// <summary>
/// Verifies that the noisy .NET built-in "System.Net.Http" meter instruments
/// (http.client.open_connections, http.client.active_requests, etc.) are dropped
/// from the OpenTelemetry metrics pipeline before export.
///
/// Background: Microsoft.ApplicationInsights.WorkerService 3.x unconditionally calls
/// meterProviderBuilder.AddMeter("System.Net.Http") with no opt-out via
/// ApplicationInsightsServiceOptions. Since ExcelMcp.McpServer makes at most one
/// lightweight HttpClient call per process (NuGetVersionChecker), these connection-pool
/// gauges/histograms were driving ~96% of billed Application Insights ingestion.
/// </summary>
[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
public sealed class ProgramTelemetryMetricsTests
{
    [Fact]
    public void ConfigureDroppedHttpClientMetrics_DropsAllNoisyHttpClientInstruments()
    {
        using var meter = new Meter(nameof(ConfigureDroppedHttpClientMetrics_DropsAllNoisyHttpClientInstruments));

        var openConnections = meter.CreateUpDownCounter<long>("http.client.open_connections");
        var activeRequests = meter.CreateUpDownCounter<long>("http.client.active_requests");
        var requestDuration = meter.CreateHistogram<double>("http.client.request.duration");
        var connectionDuration = meter.CreateHistogram<double>("http.client.connection.duration");
        var timeInQueue = meter.CreateHistogram<double>("http.client.request.time_in_queue");
        var keptMetric = meter.CreateCounter<long>("tool.invocations");

        var exportedMetrics = new List<Metric>();

        var builder = Sdk.CreateMeterProviderBuilder()
            .AddMeter(meter.Name);

        Program.ConfigureDroppedHttpClientMetrics(builder);

        using var provider = builder
            .AddInMemoryExporter(exportedMetrics)
            .Build();

        openConnections.Add(1);
        activeRequests.Add(1);
        requestDuration.Record(12.3);
        connectionDuration.Record(45.6);
        timeInQueue.Record(7.8);
        keptMetric.Add(1);

        var flushed = provider.ForceFlush();
        Assert.True(flushed, "ForceFlush should succeed; a false result would make the metric assertions below misleading.");

        var exportedNames = exportedMetrics.Select(m => m.Name).ToList();

        Assert.DoesNotContain("http.client.open_connections", exportedNames);
        Assert.DoesNotContain("http.client.active_requests", exportedNames);
        Assert.DoesNotContain("http.client.request.duration", exportedNames);
        Assert.DoesNotContain("http.client.connection.duration", exportedNames);
        Assert.DoesNotContain("http.client.request.time_in_queue", exportedNames);
        Assert.Contains("tool.invocations", exportedNames);
    }
}
