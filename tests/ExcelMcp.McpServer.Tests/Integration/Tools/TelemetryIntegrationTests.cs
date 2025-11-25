// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Diagnostics;
using OpenTelemetry;
using OpenTelemetry.Trace;
using Sbroenne.ExcelMcp.McpServer.Models;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Integration test that demonstrates telemetry output during tool invocations.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Telemetry")]
public class TelemetryIntegrationTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly TracerProvider _tracerProvider;
    private readonly List<Activity> _capturedActivities = [];

    public TelemetryIntegrationTests(ITestOutputHelper output)
    {
        _output = output;

        // Configure OpenTelemetry to capture activities for testing
        _tracerProvider = Sdk.CreateTracerProviderBuilder()
            .AddSource(ExcelMcpTelemetry.ActivitySource.Name)
            .AddProcessor(new SensitiveDataRedactingProcessor())
            .AddProcessor(new TestActivityProcessor(_capturedActivities, _output))
            .Build()!;
    }

    public void Dispose()
    {
        _tracerProvider.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ToolInvocation_TracksToTelemetry()
    {
        _output.WriteLine("=== TELEMETRY INTEGRATION TEST ===\n");

        // Act - call a tool method that uses ExecuteToolAction
        // Using Test action since it doesn't require an actual file
        var result = ExcelFileTool.ExcelFile(
            FileAction.Test,
            excelPath: "C:\\fake\\test.xlsx",
            sessionId: null);

        _output.WriteLine($"\nTool result: {result[..Math.Min(200, result.Length)]}...\n");

        // Assert - telemetry was captured
        Assert.NotEmpty(_capturedActivities);

        var activity = _capturedActivities.First();
        _output.WriteLine("=== CAPTURED TELEMETRY ===");
        _output.WriteLine($"Activity Name: {activity.DisplayName}");
        _output.WriteLine($"Duration: {activity.Duration.TotalMilliseconds:F2}ms");
        _output.WriteLine($"Status: {activity.Status}");
        _output.WriteLine("Tags:");
        foreach (var tag in activity.TagObjects)
        {
            _output.WriteLine($"  {tag.Key}: {tag.Value}");
        }
    }

    /// <summary>
    /// Simple processor that captures activities and logs them to test output.
    /// </summary>
    private sealed class TestActivityProcessor(List<Activity> activities, ITestOutputHelper output)
        : BaseProcessor<Activity>
    {
        public override void OnEnd(Activity activity)
        {
            activities.Add(activity);
            output.WriteLine($"[TELEMETRY] {activity.DisplayName} - {activity.Duration.TotalMilliseconds:F2}ms - {activity.Status}");
            base.OnEnd(activity);
        }
    }
}
