using System.IO.Pipelines;
using System.Reflection;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.WorkerService;
using Microsoft.Extensions.Configuration;
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
        // Register assembly resolver for office.dll (Microsoft.Office.Core), which is a
        // .NET Framework GAC assembly that .NET Core cannot find via standard probing.
        // office.dll is copied to our output directory by Directory.Build.targets.
        RegisterOfficeAssemblyResolver();

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

        // Disable FileSystemWatcher for config file reload.
        // Host.CreateApplicationBuilder() enables reloadOnChange:true by default, creating a
        // FileSystemWatcher for appsettings.json. Under file I/O storms (Excel temp files, lock
        // files), this watcher fires ParseEventBufferAndNotifyForEach in a tight loop on the
        // threadpool, consuming ~85% CPU. Since MCP server config never changes at runtime,
        // disable reload entirely to eliminate the watcher.
        // Re-add JSON, environment variables, and CLI args — minus the file watchers.
        builder.Configuration.Sources.Clear();
        builder.Configuration
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: false)
            .AddEnvironmentVariables()
            .AddCommandLine(args);

        // For stdio transport: Clear console logging to avoid polluting stderr with info messages.
        // The MCP client interprets stderr output as errors/warnings, so we only log Warning+
        // to stderr for debugging purposes. The MCP SDK handles protocol-level logging.
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole(consoleLogOptions =>
        {
            // Only log Warning and above to stderr - Info/Debug would appear as errors in MCP clients
            consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Warning;
        });
        builder.Logging.SetMinimumLevel(LogLevel.Warning);

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
                    1. file(action:'open') → returns sessionId
                    2. Use sessionId with ALL subsequent tools
                    3. file(action:'close', save:true/false) → ONLY when completely done

                    CALCULATION MODE:
                    - When a task mentions manual/automatic calculation or explicit recalculation, you MUST use calculation_mode.
                    - Sequence: set-mode manual → perform writes → calculate (scope: workbook) → set-mode automatic.
                    - Use get-mode when user asks for current calculation mode.

                    CRITICAL - DO NOT CLOSE SESSION PREMATURELY:
                    - Server automatically tracks active operations per session
                    - Close will be BLOCKED if operations are still running (returns error with count)
                    - Wait for error message to clear before retrying close
                    - This prevents data loss from closing mid-operation

                    SHOW EXCEL — "Agent Mode" (hero feature):
                    - Default is show:false (hidden) — but ASK the user before starting multi-step tasks
                    - Excel MCP supports "Agent Mode": users watch AI work in Excel in real-time
                    - When starting a task, present two clear action card choices:
                      (1) "Watch me work" — Show Excel side-by-side so you see every change live. Slightly slower because Excel renders each update.
                      (2) "Work in background" — Keep Excel hidden for maximum speed. You won't see changes until done, but operations complete faster.
                    - Each choice should clearly state its tradeoff so the user can make an informed decision
                    - Skip asking only when: user already stated a preference, or it's a simple one-shot operation
                    - If user picks "Watch me work": window(action:'show') + window(action:'arrange', preset:'right-half')
                    - Use window(action:'set-status-bar', text:'...') to show what you're doing in Excel's status bar
                    - Use window(action:'clear-status-bar') when done
                    - Use window(action:'hide') to hide Excel again

                    WHEN TO SKIP ASKING:
                    - User says "show me", "watch", "let me see" — show immediately, no need to ask
                    - User says "just do it", "work in background" — keep hidden, no need to ask
                    - Simple one-shot operations (read a value, check a formula) — keep hidden
                    - If user doesn't respond to the question, keep hidden

                    WHEN Excel is visible — ASK BEFORE CLOSING:
                    - If Excel is visible, the user is actively watching
                    - ALWAYS ask before closing: "Would you like me to save and close, or keep it open?"
                    - User may want to inspect results or make manual changes
                    - Do NOT auto-close visible Excel sessions
                    - Check visibility with window(action:'get-info') if unsure
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

        // Note: Update checks are handled by ExcelMCP Service (shown via Windows notification)
        // to avoid duplicate notifications when running in unified package mode

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
        finally
        {
            // CRITICAL: Auto-save all sessions and clean up Excel processes on shutdown.
            // Without this, MCP client disconnect or process exit silently discards all unsaved work.
            ServiceBridge.ServiceBridge.Dispose();
        }
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
    private static void RegisterOfficeAssemblyResolver()
    {
        AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            var name = new AssemblyName(args.Name);
            if (!string.Equals(name.Name, "office", StringComparison.OrdinalIgnoreCase))
                return null;

            var path = Path.Combine(AppContext.BaseDirectory, "office.dll");
            return File.Exists(path) ? Assembly.LoadFrom(path) : null;
        };
    }

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
        var latestVersion = await Infrastructure.McpServerVersionChecker.CheckForUpdateAsync();
        if (latestVersion != null)
        {
            Console.WriteLine();
            Console.WriteLine($"Update available: {currentVersion} -> {latestVersion}");
            Console.WriteLine("Run: dotnet tool update --global Sbroenne.ExcelMcp.McpServer");
            Console.WriteLine("Release notes: https://github.com/sbroenne/mcp-server-excel/releases/latest");
        }
    }
}



