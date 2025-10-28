using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.Core;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.McpServer.Completions;
using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>
/// ExcelCLI Model Context Protocol (MCP) Server
/// Provides resource-based tools for AI assistants to automate Excel operations:
/// - excel_file: Create, validate Excel files
/// - excel_powerquery: Manage Power Query M code and connections
/// - excel_worksheet: Worksheet lifecycle (list, create, rename, copy, delete)
/// - excel_range: Unified range operations (values, formulas, clear, copy, insert/delete, find, sort, hyperlinks)
/// - excel_parameter: Manage named ranges as parameters
/// - excel_vba: VBA script management and execution
/// - excel_connection: Manage Excel connections (OLEDB, ODBC, Text, Web)
/// - excel_data_model: Manage Data Model (tables, measures, relationships)
/// - excel_table: Manage Excel Tables (ListObjects)
/// - excel_version: Check for updates on NuGet.org
/// - begin_excel_batch/commit_excel_batch: Batch session management for multi-operation workflows
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

        // MCP Server architecture:
        // - Batch session management: LLM controls workbook lifecycle via begin/commit tools
        // - Single operations: Backward-compatible with automatic batch-of-one (when no batchId)
        // - MCP Prompts: Educate LLMs about batch workflows via [McpServerPrompt] attributes
        // - Completions: Available via ExcelCompletionHandler (manual JSON-RPC handling required)

        // Add MCP server with Excel tools (auto-discovers tools and prompts via attributes)
        builder.Services
            .AddMcpServer()
            .WithStdioServerTransport()
            .WithToolsFromAssembly();

        // Note: Completion support requires manual JSON-RPC method handling
        // See ExcelCompletionHandler for completion logic implementation
        // To enable: handle "completion/complete" method in custom transport layer

        var host = builder.Build();

        // Register cleanup handler for batch sessions
        var lifetime = host.Services.GetRequiredService<IHostApplicationLifetime>();
        lifetime.ApplicationStopping.Register(() =>
        {
            BatchSessionTool.CleanupAllBatches().GetAwaiter().GetResult();
        });

        await host.RunAsync();
    }


}
