using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.WorkerService;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.McpServer.Telemetry;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>
/// ExcelCLI Model Context Protocol (MCP) Server
/// Provides resource-based tools for AI assistants to automate Excel operations:
///
/// </summary>
public class Program
{
    public static async Task Main(string[] args)
    {
        // Register global exception handlers for unhandled exceptions (telemetry)
        RegisterGlobalExceptionHandlers();

        var builder = Host.CreateApplicationBuilder(args);

        // Configure logging to stderr for MCP protocol compliance
        builder.Logging.AddConsole(consoleLogOptions =>
        {
            consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Trace;
        });

        // Configure Application Insights Worker Service SDK for telemetry (if not opted out)
        ConfigureTelemetry(builder);

        // MCP Server architecture:
        // - Batch session management: LLM controls workbook lifecycle via begin/commit tools

        // Add MCP server with Excel tools (auto-discovers tools and prompts via attributes)
        builder.Services
            .AddMcpServer()
            .WithStdioServerTransport()
            .WithToolsFromAssembly();

        // Note: Completion support requires manual JSON-RPC method handling
        // See ExcelCompletionHandler for completion logic implementation
        // To enable: handle "completion/complete" method in custom transport layer

        var host = builder.Build();

        // Resolve TelemetryClient from DI and store for static access
        // Worker Service SDK manages the TelemetryClient lifecycle including flush on shutdown
        var telemetryClient = host.Services.GetService<TelemetryClient>();
        if (telemetryClient != null)
        {
            ExcelMcpTelemetry.SetTelemetryClient(telemetryClient);

            if (ExcelMcpTelemetry.IsDebugMode())
            {
                Console.Error.WriteLine($"[Telemetry] Application Insights configured via Worker Service SDK - User.Id={ExcelMcpTelemetry.UserId}, Session.Id={ExcelMcpTelemetry.SessionId}");
            }
        }

        // Register telemetry flush on application shutdown as backup
        // Worker Service SDK handles this automatically, but explicit flush ensures no data loss
        var lifetime = host.Services.GetService<IHostApplicationLifetime>();
        lifetime?.ApplicationStopping.Register(() =>
        {
            ExcelMcpTelemetry.Flush();
        });

        await host.RunAsync();
    }

    /// <summary>
    /// Configures Application Insights Worker Service SDK for telemetry.
    /// Uses AddApplicationInsightsTelemetryWorkerService() for proper host integration.
    /// Enables Users/Sessions/Funnels/User Flows analytics in Azure Portal.
    /// </summary>
    private static void ConfigureTelemetry(HostApplicationBuilder builder)
    {
        // Debug mode: log telemetry to stderr for local testing
        var isDebugMode = ExcelMcpTelemetry.IsDebugMode();

        var connectionString = ExcelMcpTelemetry.GetConnectionString();
        if (string.IsNullOrEmpty(connectionString) && !isDebugMode)
        {
            return; // No connection string available and not in debug mode
        }

        if (isDebugMode && string.IsNullOrEmpty(connectionString))
        {
            // Debug mode without connection string - telemetry will be tracked but not sent
            Console.Error.WriteLine("[Telemetry] Debug mode enabled - telemetry tracked locally (no Azure connection)");
        }

        // Configure Application Insights Worker Service SDK
        // This provides:
        // - Proper DI integration with IHostApplicationLifetime
        // - Automatic dependency tracking
        // - Automatic performance counter collection (where available)
        // - Proper telemetry channel with ServerTelemetryChannel (retries, local storage)
        // - Automatic flush on host shutdown
        var aiOptions = new ApplicationInsightsServiceOptions
        {
            // Set connection string if available
            ConnectionString = connectionString,

            // Disable features not needed for MCP server (reduces overhead)
            EnableHeartbeat = true,  // Useful for monitoring server health
            EnableAdaptiveSampling = true,  // Helps manage telemetry volume
            EnableQuickPulseMetricStream = false,  // Live Metrics not needed for CLI tool
            EnablePerformanceCounterCollectionModule = false,  // Perf counters not useful for short-lived CLI
            EnableEventCounterCollectionModule = false,  // Event counters not needed

            // Enable dependency tracking for HTTP calls
            EnableDependencyTrackingTelemetryModule = true,
        };

        builder.Services.AddApplicationInsightsTelemetryWorkerService(aiOptions);

        // Add custom telemetry initializer for User.Id and Session.Id
        // This enables the Users and Sessions blades in Azure Portal
        builder.Services.AddSingleton<Microsoft.ApplicationInsights.Extensibility.ITelemetryInitializer, ExcelMcpTelemetryInitializer>();
    }

    /// <summary>
    /// Registers global exception handlers to capture unhandled exceptions.
    /// </summary>
    private static void RegisterGlobalExceptionHandlers()
    {
        // Handle exceptions that escape all catch blocks
        AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
        {
            if (e.ExceptionObject is Exception ex)
            {
                ExcelMcpTelemetry.TrackUnhandledException(ex, "AppDomain.UnhandledException");
            }
        };

        // Handle unobserved task exceptions
        TaskScheduler.UnobservedTaskException += (sender, e) =>
        {
            ExcelMcpTelemetry.TrackUnhandledException(e.Exception, "TaskScheduler.UnobservedTaskException");
            // Don't observe it - let the runtime handle it
        };
    }
}

