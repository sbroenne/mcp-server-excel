using Microsoft.ApplicationInsights.WorkerService;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Console;
using Microsoft.Extensions.Options;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Collection("ProgramTransport")]
[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "ProgramTransport")]
[Trait("Speed", "Fast")]
public sealed class ProgramLoggingTests
{
    [Fact]
    public void ConfigureStdioLogging_RoutesAllConsoleLogsToStandardError()
    {
        var services = new ServiceCollection();

        services.AddLogging(Program.ConfigureStdioLogging);

        using var provider = services.BuildServiceProvider();
        var options = provider.GetRequiredService<IOptions<ConsoleLoggerOptions>>().Value;

        Assert.Equal(LogLevel.Trace, options.LogToStandardErrorThreshold);
    }

    [Fact]
    public void ConfigureStdioLogging_DoesNotWriteInfoLogsToStandardOutputWhenMinimumLevelIsOverridden()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var originalOut = Console.Out;
        var originalError = Console.Error;

        try
        {
            Console.SetOut(stdout);
            Console.SetError(stderr);

            var services = new ServiceCollection();
            services.AddLogging(builder =>
            {
                Program.ConfigureStdioLogging(builder);
                builder.SetMinimumLevel(LogLevel.Information);
            });

            using (var provider = services.BuildServiceProvider())
            {
                var logger = provider.GetRequiredService<ILoggerFactory>().CreateLogger("Microsoft.Hosting.Lifetime");
                logger.LogInformation("host started");
            }
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }

        Assert.Empty(stdout.ToString());
        Assert.Contains("host started", stderr.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void ConfigureStdioLogging_SuppressesApplicationInsightsInfoLogs()
    {
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        var originalOut = Console.Out;
        var originalError = Console.Error;

        try
        {
            Console.SetOut(stdout);
            Console.SetError(stderr);

            var services = new ServiceCollection();
            services.AddLogging(Program.ConfigureStdioLogging);
            services.AddApplicationInsightsTelemetryWorkerService(new ApplicationInsightsServiceOptions
            {
                ConnectionString = "InstrumentationKey=00000000-0000-0000-0000-000000000000"
            });

            using (var provider = services.BuildServiceProvider())
            {
                var logger = provider.GetRequiredService<ILoggerFactory>().CreateLogger("Microsoft.ApplicationInsights.TelemetryClient");
                logger.LogInformation("telemetry configured");
                logger.LogWarning("telemetry warning");
            }
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }

        Assert.Empty(stdout.ToString());
        Assert.DoesNotContain("telemetry configured", stderr.ToString(), StringComparison.Ordinal);
        Assert.Contains("telemetry warning", stderr.ToString(), StringComparison.Ordinal);
    }
}
