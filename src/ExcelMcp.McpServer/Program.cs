using Azure.Monitor.OpenTelemetry.Exporter;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OpenTelemetry.Trace;
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

        // Configure OpenTelemetry for Application Insights (if not opted out)
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

        await host.RunAsync();
    }

    /// <summary>
    /// Configures OpenTelemetry with Azure Monitor exporter for Application Insights.
    /// Respects opt-out via EXCELMCP_TELEMETRY_OPTOUT environment variable.
    /// </summary>
    private static void ConfigureTelemetry(HostApplicationBuilder builder)
    {
        // Check if telemetry is enabled
        if (ExcelMcpTelemetry.IsOptedOut())
        {
            return; // User opted out
        }

        // Debug mode: log telemetry to stderr instead of Azure (for local testing)
        var isDebugMode = ExcelMcpTelemetry.IsDebugMode();

        var connectionString = ExcelMcpTelemetry.GetConnectionString();
        if (string.IsNullOrEmpty(connectionString) && !isDebugMode)
        {
            return; // No connection string available and not in debug mode
        }

        // Configure OpenTelemetry
        builder.Services.AddOpenTelemetry()
            .WithTracing(tracing =>
            {
                tracing
                    .AddSource(ExcelMcpTelemetry.ActivitySource.Name)
                    .AddProcessor(new SensitiveDataRedactingProcessor());

                if (isDebugMode)
                {
                    // Debug mode: write to stderr (console) for local testing
                    tracing.AddConsoleExporter();
                    Console.Error.WriteLine("[Telemetry] Debug mode enabled - logging to stderr");
                }
                else
                {
                    // Production: send to Azure Monitor
                    tracing.AddAzureMonitorTraceExporter(options =>
                    {
                        options.ConnectionString = connectionString;
                    });
                }
            });
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

