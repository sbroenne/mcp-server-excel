using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

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

        // Add MCP server with Excel tools
        builder.Services
            .AddMcpServer()
            .WithStdioServerTransport()
            .WithToolsFromAssembly();

        await builder.Build().RunAsync();
    }


}
