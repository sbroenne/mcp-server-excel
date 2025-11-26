using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
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
    /// Configures Application Insights SDK for telemetry.
    /// Enables Users/Sessions/Funnels/User Flows analytics in Azure Portal.
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

        // Configure Application Insights SDK
        var aiConfig = TelemetryConfiguration.CreateDefault();
        if (!string.IsNullOrEmpty(connectionString))
        {
            aiConfig.ConnectionString = connectionString;
        }
        else if (isDebugMode)
        {
            // Debug mode without connection string - telemetry will be tracked but not sent
            // This allows testing the tracking code without Azure resources
            Console.Error.WriteLine("[Telemetry] Debug mode enabled - telemetry tracked locally (no Azure connection)");
        }

        // Add initializer to set User.Id and Session.Id on all telemetry
        aiConfig.TelemetryInitializers.Add(new ExcelMcpTelemetryInitializer());

        // Register TelemetryClient as singleton for dependency injection
        var telemetryClient = new TelemetryClient(aiConfig);
        builder.Services.AddSingleton(telemetryClient);
        builder.Services.AddSingleton(aiConfig);

        // Store reference for static access in ExcelMcpTelemetry
        ExcelMcpTelemetry.SetTelemetryClient(telemetryClient);

        if (isDebugMode)
        {
            Console.Error.WriteLine($"[Telemetry] Application Insights configured - User.Id={ExcelMcpTelemetry.UserId}, Session.Id={ExcelMcpTelemetry.SessionId}");
        }
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

