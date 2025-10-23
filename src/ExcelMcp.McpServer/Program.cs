using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.Core;
using Sbroenne.ExcelMcp.McpServer.Tools;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>
/// ExcelCLI Model Context Protocol (MCP) Server
/// Provides 6 resource-based tools for AI assistants to automate Excel operations:
/// - excel_file: Create, validate Excel files
/// - excel_powerquery: Manage Power Query M code and connections
/// - excel_worksheet: CRUD operations on worksheets and data
/// - excel_parameter: Manage named ranges as parameters
/// - excel_cell: Individual cell operations (get/set values/formulas)
/// - excel_vba: VBA script management and execution
/// - excel_version: Check for updates on NuGet.org
///
/// Performance Optimization:
/// Uses ExcelInstancePool for conversational workflows - reuses Excel instances
/// across multiple operations on the same workbook, reducing startup overhead
/// from ~2-5 seconds per operation to near-instantaneous for cached instances.
/// </summary>
public class Program
{
    public static async Task Main(string[] args)
    {
        var builder = Host.CreateApplicationBuilder(args);

        // Configure logging to stderr for MCP protocol compliance
        builder.Logging.AddConsole(consoleLogOptions =>
        {
            consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Trace;
        });

        // Check for updates on startup (non-blocking)
        _ = Task.Run(async () =>
        {
            try
            {
                var checker = new VersionChecker();
                var result = await checker.CheckForUpdatesAsync("Sbroenne.ExcelMcp.McpServer");
                
                if (result.Success && result.IsOutdated)
                {
                    // Log warning to stderr for MCP protocol compliance
                    Console.Error.WriteLine($"⚠️  WARNING: ExcelMcp update available!");
                    Console.Error.WriteLine($"   Current version: {result.CurrentVersion}");
                    Console.Error.WriteLine($"   Latest version:  {result.LatestVersion}");
                    Console.Error.WriteLine($"   The dnx command automatically downloads the latest version.");
                    Console.Error.WriteLine($"   Restart VS Code to update to the new version.");
                }
            }
            catch
            {
                // Silently ignore version check failures on startup
            }
        });

        // Register Excel instance pool as singleton for reuse across tool calls
        // Idle instances are automatically cleaned up after 60 seconds
        builder.Services.AddSingleton<ExcelInstancePool>(sp =>
            new ExcelInstancePool(idleTimeout: TimeSpan.FromSeconds(60)));

        // Initialize the pool for use by Core commands and MCP tools
        var pool = new ExcelInstancePool(idleTimeout: TimeSpan.FromSeconds(60));

        // Configure Core layer to use pooling (zero-change integration)
        ExcelHelper.InstancePool = pool;

        // Configure MCP tools layer to use pooling (for static access)
        ExcelToolsPoolManager.Initialize(pool);

        // Add MCP server with Excel tools
        builder.Services
            .AddMcpServer()
            .WithStdioServerTransport()
            .WithToolsFromAssembly();

        var host = builder.Build();

        // Ensure pool is disposed on shutdown
        var lifetime = host.Services.GetRequiredService<IHostApplicationLifetime>();
        lifetime.ApplicationStopping.Register(() =>
        {
            // Clear pool references
            ExcelHelper.InstancePool = null;
            ExcelToolsPoolManager.Shutdown();

            // Dispose pool instance
            pool.Dispose();
        });

        await host.RunAsync();
    }


}
