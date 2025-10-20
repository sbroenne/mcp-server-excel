using Xunit;
using Xunit.Abstractions;
using System.Text.Json;
using System.Diagnostics;
using System.Text;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Test to diagnose MCP Server framework parameter binding issues
/// by testing with minimal validation attributes
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
public class McpParameterBindingTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private Process? _serverProcess;

    public McpParameterBindingTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"MCPBinding_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (_serverProcess != null)
        {
            try 
            {
                if (!_serverProcess.HasExited)
                {
                    _serverProcess.Kill();
                }
            }
            catch (Exception)
            {
                // Process cleanup error - ignore
            }
        }
        _serverProcess?.Dispose();
        
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch
        {
            // Cleanup failed - not critical
        }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task McpServer_BasicParameterBinding_ShouldWork()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);
        
        var testFile = Path.Combine(_tempDir, "binding-test.xlsx");

        // Act & Assert
        _output.WriteLine("=== MCP Parameter Binding Test ===");
        
        // First, let's see what tools are available
        _output.WriteLine("Querying available tools...");
        var toolsListRequest = new
        {
            jsonrpc = "2.0",
            id = Environment.TickCount,
            method = "tools/list",
            @params = new { }
        };
        
        var toolsListJson = JsonSerializer.Serialize(toolsListRequest);
        _output.WriteLine($"Sending tools list: {toolsListJson}");
        await server.StandardInput.WriteLineAsync(toolsListJson);
        await server.StandardInput.FlushAsync();
        
        var toolsListResponse = await server.StandardOutput.ReadLineAsync();
        _output.WriteLine($"Available tools: {toolsListResponse}");
        
        // Test the original excel_file tool to see what specific error occurs
        _output.WriteLine("Testing excel_file tool through MCP framework...");
        var response = await CallExcelTool(server, "excel_file", new 
        { 
            action = "create-empty", 
            excelPath = testFile
        });
        
        _output.WriteLine($"MCP Response: {response}");
        
        // Parse response to understand what happened
        var jsonDoc = JsonDocument.Parse(response);
        
        // Handle different response formats
        if (jsonDoc.RootElement.TryGetProperty("error", out var errorProperty))
        {
            // Standard JSON-RPC error
            var code = errorProperty.GetProperty("code").GetInt32();
            var message = errorProperty.GetProperty("message").GetString();
            _output.WriteLine($"❌ JSON-RPC Error {code}: {message}");
            Assert.Fail($"JSON-RPC error {code}: {message}");
        }
        else if (jsonDoc.RootElement.TryGetProperty("result", out var result))
        {
            if (result.TryGetProperty("isError", out var isErrorElement) && isErrorElement.GetBoolean())
            {
                var errorContent = result.GetProperty("content")[0].GetProperty("text").GetString();
                _output.WriteLine($"❌ MCP Framework Error: {errorContent}");
                
                // This is the key error we're trying to debug
                _output.WriteLine("🔍 This confirms the MCP framework is catching and suppressing the actual error");
                Assert.Fail($"MCP framework error: {errorContent}");
            }
            else
            {
                var contentText = result.GetProperty("content")[0].GetProperty("text").GetString();
                _output.WriteLine($"✅ MCP Success: {contentText}");
                
                // Parse the tool response
                var toolResult = JsonDocument.Parse(contentText!);
                if (toolResult.RootElement.TryGetProperty("success", out var successElement))
                {
                    var success = successElement.GetBoolean();
                    Assert.True(success, $"Tool execution failed: {contentText}");
                    Assert.True(File.Exists(testFile), "File was not created");
                }
                else
                {
                    Assert.Fail($"Unexpected tool response format: {contentText}");
                }
            }
        }
        else
        {
            Assert.Fail($"Unexpected response format: {response}");
        }
    }

    private Process StartMcpServer()
    {
        // Find the workspace root directory
        var currentDir = Directory.GetCurrentDirectory();
        var workspaceRoot = currentDir;
        while (!File.Exists(Path.Combine(workspaceRoot, "Sbroenne.ExcelMcp.sln")))
        {
            var parent = Directory.GetParent(workspaceRoot);
            if (parent == null) break;
            workspaceRoot = parent.FullName;
        }
        
        var serverPath = Path.Combine(workspaceRoot, "src", "ExcelMcp.McpServer", "bin", "Debug", "net9.0", "Sbroenne.ExcelMcp.McpServer.exe");
        _output.WriteLine($"Looking for server at: {serverPath}");
        
        if (!File.Exists(serverPath))
        {
            _output.WriteLine("Server not found, building first...");
            // Try to build first
            var buildProcess = Process.Start(new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = "build src/ExcelMcp.McpServer/ExcelMcp.McpServer.csproj",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = workspaceRoot
            });
            buildProcess!.WaitForExit();
            _output.WriteLine($"Build exit code: {buildProcess.ExitCode}");
        }

        var server = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = serverPath,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };

        server.Start();
        _serverProcess = server;
        return server;
    }

    private async Task InitializeServer(Process server)
    {
        var initRequest = new
        {
            jsonrpc = "2.0",
            id = 1,
            method = "initialize",
            @params = new
            {
                protocolVersion = "2024-11-05",
                capabilities = new { },
                clientInfo = new
                {
                    name = "Test",
                    version = "1.0.0"
                }
            }
        };

        var json = JsonSerializer.Serialize(initRequest);
        _output.WriteLine($"Sending init: {json}");
        
        await server.StandardInput.WriteLineAsync(json);
        await server.StandardInput.FlushAsync();

        // Read and verify response
        var response = await server.StandardOutput.ReadLineAsync();
        _output.WriteLine($"Received init response: {response}");
        Assert.NotNull(response);
    }

    private async Task<string> CallExcelTool(Process server, string toolName, object arguments)
    {
        var request = new
        {
            jsonrpc = "2.0",
            id = Environment.TickCount,
            method = "tools/call",
            @params = new
            {
                name = toolName,
                arguments = arguments
            }
        };

        var json = JsonSerializer.Serialize(request);
        _output.WriteLine($"Sending tool call: {json}");
        
        await server.StandardInput.WriteLineAsync(json);
        await server.StandardInput.FlushAsync();

        var response = await server.StandardOutput.ReadLineAsync();
        _output.WriteLine($"Received tool response: {response}");
        Assert.NotNull(response);
        
        return response;
    }
}
