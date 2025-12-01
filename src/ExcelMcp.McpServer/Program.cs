using System.IO.Pipelines;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.WorkerService;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.McpServer.Telemetry;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>
/// ExcelMCP Model Context Protocol (MCP) Server.
/// Provides resource-based tools for AI assistants to automate Excel operations.
/// </summary>
public class Program
{
    // Test transport configuration - set by tests before calling Main()
    // These are intentionally static for test injection. Thread-safety is not required
    // because tests run sequentially and call ResetTestTransport() after each test.
    private static Pipe? _testInputPipe;
    private static Pipe? _testOutputPipe;

    /// <summary>
    /// Configures the server to use in-memory pipe transport for testing.
    /// Call this before RunAsync() to enable test mode.
    /// </summary>
    /// <param name="inputPipe">Pipe for reading client requests (client writes, server reads)</param>
    /// <param name="outputPipe">Pipe for writing server responses (server writes, client reads)</param>
    public static void ConfigureTestTransport(Pipe inputPipe, Pipe outputPipe)
    {
        _testInputPipe = inputPipe;
        _testOutputPipe = outputPipe;
    }

    /// <summary>
    /// Resets test transport configuration (call after test completes).
    /// </summary>
    public static void ResetTestTransport()
    {
        _testInputPipe = null;
        _testOutputPipe = null;
    }

    public static async Task<int> Main(string[] args)
    {
        // Register global exception handlers for unhandled exceptions (telemetry)
        RegisterGlobalExceptionHandlers();

        var builder = Host.CreateApplicationBuilder(args);

        // Configure logging to stderr for MCP protocol compliance
        builder.Logging.AddConsole(consoleLogOptions =>
        {
            consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Trace;
        });

        // Configure Application Insights
        ConfigureTelemetry(builder);

        // Configure MCP Server - use test transport if configured, otherwise stdio
        var mcpBuilder = builder.Services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new()
                {
                    Name = "excel-mcp",
                    Version = typeof(Program).Assembly.GetName().Version?.ToString() ?? "1.0.0"
                };

                // Server-wide instructions for LLMs - helps with tool selection and workflow understanding
                options.ServerInstructions = """
                    ExcelMCP automates Microsoft Excel via COM interop.

                    CRITICAL: File must be CLOSED in Excel desktop app (COM requires exclusive access).

                    SESSION LIFECYCLE:
                    1. excel_file(action:'open') → returns sessionId
                    2. Use sessionId with ALL subsequent tools
                    3. excel_file(action:'close', save:true/false) → ONLY when completely done
                    Keep session open across multiple operations - don't close prematurely!
                    """;
            })
            .WithToolsFromAssembly()
            .WithPromptsFromAssembly(); // Auto-discover prompts marked with [McpServerPromptType]

        if (_testInputPipe != null && _testOutputPipe != null)
        {
            // Test mode: use in-memory pipe transport
            mcpBuilder.WithStreamServerTransport(
                _testInputPipe.Reader.AsStream(),
                _testOutputPipe.Writer.AsStream());
        }
        else
        {
            // Production mode: use stdio transport
            mcpBuilder.WithStdioServerTransport();
        }

        var host = builder.Build();

        // Initialize telemetry client for static access
        InitializeTelemetryClient(host.Services);

        await host.RunAsync();

        return 0;
    }

    /// <summary>
    /// Initializes the static TelemetryClient from DI container.
    /// </summary>
    private static void InitializeTelemetryClient(IServiceProvider services)
    {
        // Resolve TelemetryClient from DI and store for static access
        // Worker Service SDK manages the TelemetryClient lifecycle including flush on shutdown
        var telemetryClient = services.GetService<TelemetryClient>();
        if (telemetryClient != null)
        {
            ExcelMcpTelemetry.SetTelemetryClient(telemetryClient);

            if (ExcelMcpTelemetry.IsDebugMode())
            {
                Console.Error.WriteLine($"[Telemetry] Application Insights configured via Worker Service SDK - User.Id={ExcelMcpTelemetry.UserId}, Session.Id={ExcelMcpTelemetry.SessionId}");
            }
        }
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

            // Disable dependency tracking for HTTP calls
            EnableDependencyTrackingTelemetryModule = false,
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

