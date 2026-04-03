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
    private static readonly object TestTransportLock = new();

    // Test transport configuration - set by tests before calling Main()
    // These are intentionally static for test injection, but we still guard them so
    // leaked test state fails fast instead of contaminating the next transport-backed test.
    private static Pipe? _testInputPipe;
    private static Pipe? _testOutputPipe;
    private static CancellationTokenSource? _testShutdownCts;
    private static long _testTransportGeneration;

    /// <summary>
    /// Configures the server to use in-memory pipe transport for testing.
    /// Call this before RunAsync() to enable test mode.
    /// </summary>
    /// <param name="inputPipe">Pipe for reading client requests (client writes, server reads)</param>
    /// <param name="outputPipe">Pipe for writing server responses (server writes, client reads)</param>
    public static void ConfigureTestTransport(Pipe inputPipe, Pipe outputPipe)
    {
        lock (TestTransportLock)
        {
            if (_testInputPipe != null || _testOutputPipe != null || _testShutdownCts != null)
            {
                throw new InvalidOperationException(
                    "Test transport is already configured. Ensure the previous MCP transport test completed cleanup before starting another one.");
            }

            _testInputPipe = inputPipe;
            _testOutputPipe = outputPipe;
            _testShutdownCts = new CancellationTokenSource();
            _testTransportGeneration++;
        }
    }

    /// <summary>
    /// Requests shutdown for the active in-memory test transport without clearing transport state.
    /// </summary>
    public static void RequestTestTransportShutdown()
    {
        CancellationTokenSource? shutdownCts;

        lock (TestTransportLock)
        {
            shutdownCts = _testShutdownCts;
        }

        if (shutdownCts != null)
        {
            try
            {
                shutdownCts.Cancel();
            }
            catch (ObjectDisposedException)
            {
                // ResetTestTransport() owns disposing the test CTS after the host has fully stopped.
            }
        }
    }

    /// <summary>
    /// Resets test transport configuration after the in-memory test host has stopped.
    /// </summary>
    public static void ResetTestTransport()
    {
        CancellationTokenSource? shutdownCts;

        lock (TestTransportLock)
        {
            shutdownCts = _testShutdownCts;
            _testShutdownCts = null;
            _testInputPipe = null;
            _testOutputPipe = null;
        }

        shutdownCts?.Dispose();

        ServiceBridge.ServiceBridge.ResetForTests();
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

        Pipe? testInputPipe;
        Pipe? testOutputPipe;
        CancellationTokenSource? testShutdownCts;
        long testTransportGeneration;

        lock (TestTransportLock)
        {
            testInputPipe = _testInputPipe;
            testOutputPipe = _testOutputPipe;
            testShutdownCts = _testShutdownCts;
            testTransportGeneration = testInputPipe != null && testOutputPipe != null
                ? _testTransportGeneration
                : 0;
        }

        if (testTransportGeneration != 0)
        {
            ServiceBridge.ServiceBridge.SetTestOwnerToken(testTransportGeneration);
        }

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

                    CALCULATION MODE (Performance Optimization):
                    - Use calculation_mode for bulk write operations (10+ cells with values or formulas).
                    - Workflow: set-mode(manual) → perform all writes → calculate(scope: workbook) → set-mode(automatic).
                    - Skips recalculation after every cell write, calculates once at end — much faster for batch operations.
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

        if (testInputPipe != null && testOutputPipe != null)
        {
            // Test mode: use in-memory pipe transport
            mcpBuilder.WithStreamServerTransport(
                testInputPipe.Reader.AsStream(),
                testOutputPipe.Writer.AsStream());
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

        var runToken = testShutdownCts?.Token ?? CancellationToken.None;

        try
        {
            await host.RunAsync(runToken);
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
            if (testTransportGeneration == 0)
            {
                ServiceBridge.ServiceBridge.Dispose();
            }
            else
            {
                ServiceBridge.ServiceBridge.DisposeIfOwnedBy(testTransportGeneration);
            }
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

            return ResolveOfficeDll();
        };
    }

    /// <summary>
    /// Resolves office.dll (Microsoft.Office.Core) from multiple locations.
    /// office.dll is a .NET Framework GAC assembly that .NET Core cannot find automatically.
    /// It is present when Microsoft Office is installed, but not in the .NET Core probing paths.
    /// Search order:
    ///   1. AppContext.BaseDirectory (copied by Directory.Build.targets in local dev builds)
    ///   2. .NET Framework GAC - v16 then v15 (v15 is accepted by the CLR for v16 requests)
    ///   3. Office installation directory (click-to-run Office 365 doesn't register in GAC)
    /// </summary>
    private static Assembly? ResolveOfficeDll()
    {
        // 1. Local build output (Directory.Build.targets copies office.dll here in dev builds)
        var localPath = Path.Combine(AppContext.BaseDirectory, "office.dll");
        if (File.Exists(localPath))
            return Assembly.LoadFrom(localPath);

        // 2. .NET Framework GAC — v16 preferred, v15 accepted (CLR honours AssemblyResolve return regardless of version)
        string[] gacPaths =
        [
            @"C:\Windows\assembly\GAC_MSIL\office\16.0.0.0__71e9bce111e9429c\OFFICE.DLL",
            @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL",
        ];
        foreach (var gacPath in gacPaths)
        {
            if (File.Exists(gacPath))
                return Assembly.LoadFrom(gacPath);
        }

        // 3. Office 365 click-to-run installation directories (Office registers its own copy)
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        string[] officeDirs =
        [
            Path.Combine(programFiles, @"Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-inv16"),
            Path.Combine(programFiles, @"Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in"),
            Path.Combine(programFilesX86, @"Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-inv16"),
            Path.Combine(programFilesX86, @"Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in"),
        ];
        foreach (var dir in officeDirs)
        {
            var officePath = Path.Combine(dir, "OFFICE.dll");
            if (File.Exists(officePath))
                return Assembly.LoadFrom(officePath);
        }

        return null;
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
              mcp-excel.exe [options]

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
            Console.WriteLine("Download: https://github.com/sbroenne/mcp-server-excel/releases/latest");
        }
    }
}



