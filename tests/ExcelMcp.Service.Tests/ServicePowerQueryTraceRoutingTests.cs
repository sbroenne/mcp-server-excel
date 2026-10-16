using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ServicePowerQueryTraceRoutingTests
{
    [Fact]
    public async Task SettingsGroup_RoutesTracingActionsWithoutWorkbookSession()
    {
        string traceFolder = Path.Combine(Path.GetTempPath(), $"service-pqtrace-{Guid.NewGuid():N}");
        Directory.CreateDirectory(traceFolder);
        try
        {
            var settings = new FakeTraceSettings(traceFolder, new PowerQueryTracingSnapshot(0, 15));
            using var service = new ExcelMcpService(null, settings);

            var readResponse = await service.ProcessAsync(new ServiceRequest
            {
                Command = "settings.get-pq-tracing"
            });
            Assert.True(readResponse.Success, readResponse.ErrorMessage);
            var readResult = JsonSerializer.Deserialize<PowerQueryTracingStatusResult>(
                readResponse.Result!,
                ServiceProtocol.JsonOptions);
            Assert.NotNull(readResult);
            Assert.False(readResult.Enabled);
            Assert.Equal(15, readResult.Level);

            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "settings.set-pq-tracing",
                Args = """{"enabled":true,"level":8}"""
            });

            Assert.True(response.Success, response.ErrorMessage);
            var result = JsonSerializer.Deserialize<PowerQueryTracingStatusResult>(
                response.Result!,
                ServiceProtocol.JsonOptions);
            Assert.NotNull(result);
            Assert.True(result.Enabled);
            Assert.Equal(8, result.Level);
            Assert.False(result.PreviousEnabled);
            Assert.Equal(15, result.PreviousLevel);
            Assert.Equal(0, service.SessionCount);
            Assert.Contains("Restart Excel", result.Note, StringComparison.Ordinal);
            service.Dispose();
            Assert.True(settings.Read().Enabled);
        }
        finally
        {
            Directory.Delete(traceFolder, recursive: true);
        }
    }

    private sealed class FakeTraceSettings(string traceFolder, PowerQueryTracingSnapshot initial) : IPowerQueryTraceSettings
    {
        private int? _enabled = initial.EnabledValue;
        private int? _level = initial.LevelValue;

        public string TraceFolder { get; } = traceFolder;

        public PowerQueryTracingSnapshot Read() => new(_enabled, _level);

        public void Write(bool enabled, int? level)
        {
            _enabled = enabled ? 1 : 0;
            _level = level ?? _level;
        }

    }
}
