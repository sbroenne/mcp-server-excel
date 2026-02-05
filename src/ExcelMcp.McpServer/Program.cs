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
        // Handle --help and --version flags for easy verification
        if (args.Length > 0)
        {
            var arg = args[0].ToLowerInvariant();
            if (arg is "-h" or "--help" or "-?" or "/?" or "/h")
            {
                ShowHelp();
                return 0;
            }
            if (arg is "-v" or "--version")
            {
                await ShowVersionAsync();
                return 0;
            }
        }

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

                    CALCULATION MODE:
                    - When a task mentions manual/automatic calculation or explicit recalculation, you MUST use excel_calculation_mode.
                    - Sequence: set-mode manual → perform writes → calculate (scope: workbook) → set-mode automatic.
                    - Use get-mode when user asks for current calculation mode.

                    CRITICAL - DO NOT CLOSE SESSION PREMATURELY:
                    - Server automatically tracks active operations per session
                    - Close will be BLOCKED if operations are still running (returns error with count)
                    - Wait for error message to clear before retrying close
                    - This prevents data loss from closing mid-operation

                    SHOW EXCEL (watch changes live):
                    - Default is showExcel:false (hidden) - USE THIS DEFAULT unless user explicitly requests visible Excel
                    - showExcel:true displays Excel window so user can watch operations in real-time
                    - showExcel:true is SLOWER - only use when user explicitly wants to watch

                    WHEN TO ASK (not auto-enable) about showExcel:
                    - Only ASK if user seems confused about what's happening
                    - Only ASK if debugging or troubleshooting issues
                    - Example question: "Would you like me to show Excel so you can watch the changes?"
                    - DO NOT default to showExcel:true - always use showExcel:false unless user says yes

                    WHEN showExcel=true - ASK BEFORE CLOSING:
                    - If Excel is visible, the user is actively watching
                    - ALWAYS ask user before closing: "Would you like me to save and close the file, or keep it open?"
                    - User may want to inspect results, make manual changes, or continue working
                    - Do NOT auto-close visible Excel sessions
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

        // Check for updates on startup (non-blocking, logs to stderr)
        _ = Task.Run(async () =>
        {
            try
            {
                // Wait briefly to avoid interfering with startup
                await Task.Delay(TimeSpan.FromSeconds(2));

                var updateInfo = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync();
                if (updateInfo != null)
                {
                    // Log to stderr directly - this is a non-critical notification
                    // Using Console.Error avoids CA1848 analyzer warning about LoggerMessage delegates
                    await Console.Error.WriteLineAsync(
                        $"[Info] MCP Server update available: {updateInfo.CurrentVersion} -> {updateInfo.LatestVersion}. " +
                        "Run: dotnet tool update --global Sbroenne.ExcelMcp.McpServer");
                }
            }
            catch
            {
                // Fail silently - version check should never interfere with server operation
            }
        });

        try
        {
            await host.RunAsync();
            return 0;
        }
        catch (OperationCanceledException)
        {
            // Graceful shutdown via cancellation (e.g., Ctrl+C, SIGTERM)
            // This is expected behavior, not an error
            return 0;
        }
#pragma warning disable CA1031 // Catch general exception - this is a top-level handler that must not crash
        catch (Exception ex)
        {
            // Track MCP SDK/transport errors (protocol errors, serialization errors, etc.)
            ExcelMcpTelemetry.TrackUnhandledException(ex, "McpServer.RunAsync");
            ExcelMcpTelemetry.Flush(); // Ensure telemetry is sent before exit

            // Return exit code 1 for fatal errors (FR-024, SC-015a)
            // Do NOT re-throw - deterministic exit code is more important for callers
            return 1;
        }
#pragma warning restore CA1031
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
        }
    }

    /// <summary>
    /// Configures Application Insights Worker Service SDK for telemetry.
    /// Uses AddApplicationInsightsTelemetryWorkerService() for proper host integration.
    /// Enables Users/Sessions/Funnels/User Flows analytics in Azure Portal.
    /// </summary>
    private static void ConfigureTelemetry(HostApplicationBuilder builder)
    {
        var connectionString = ExcelMcpTelemetry.GetConnectionString();
        if (string.IsNullOrEmpty(connectionString))
        {
            return; // No connection string available (local dev build)
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

    /// <summary>
    /// Shows help information.
    /// </summary>
    private static void ShowHelp()
    {
        var version = typeof(Program).Assembly.GetName().Version?.ToString() ?? "1.0.0";
        Console.WriteLine($"""
            Excel MCP Server v{version}

            An MCP (Model Context Protocol) server for Microsoft Excel automation.
            Provides 22 tools with 195+ operations for AI assistants.

            Usage:
              Sbroenne.ExcelMcp.McpServer.exe [options]

            Options:
              -h, --help      Show this help message
              -v, --version   Show version information

            Without options, starts the MCP server in stdio mode.

            Requirements:
              - Windows x64
              - Microsoft Excel 2016 or later (desktop version)

            Documentation:
              https://sbroenne.github.io/mcp-server-excel/

            Source:
              https://github.com/sbroenne/mcp-server-excel
            """);
    }

    /// <summary>
    /// Shows version information and checks for updates.
    /// </summary>
    private static async Task ShowVersionAsync()
    {
        var currentVersion = Infrastructure.McpServerVersionChecker.GetCurrentVersion();
        Console.WriteLine($"Excel MCP Server v{currentVersion}");

        // Check for updates (non-blocking, 5-second timeout)
        var updateInfo = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync();
        if (updateInfo != null)
        {
            Console.WriteLine();
            Console.WriteLine($"Update available: {updateInfo.CurrentVersion} -> {updateInfo.LatestVersion}");
            Console.WriteLine("Run: dotnet tool update --global Sbroenne.ExcelMcp.McpServer");
            Console.WriteLine("Release notes: https://github.com/sbroenne/mcp-server-excel/releases/latest");
        }
        else
        {
            // Check completed but no update available (or check failed silently)
            // Don't show anything - keep output clean for scripting
        }
    }
}

