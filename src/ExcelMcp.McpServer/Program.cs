using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

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
        var builder = Host.CreateApplicationBuilder(args);

        // Configure logging to stderr for MCP protocol compliance
        builder.Logging.AddConsole(consoleLogOptions =>
        {
            consoleLogOptions.LogToStandardErrorThreshold = LogLevel.Trace;
        });

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


}
