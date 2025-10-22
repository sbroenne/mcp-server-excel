using System.Diagnostics;
using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration;

/// <summary>
/// True MCP integration tests that act as MCP clients
/// These tests start the MCP server process and communicate via stdio using the MCP protocol
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "MCPProtocol")]
public class McpClientIntegrationTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private Process? _serverProcess;

    public McpClientIntegrationTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"MCPClient_Tests_{Guid.NewGuid():N}");
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
            catch (InvalidOperationException)
            {
                // Process already exited or disposed - this is fine
            }
            catch (Exception)
            {
                // Any other process cleanup error - ignore
            }
        }
        _serverProcess?.Dispose();

        if (Directory.Exists(_tempDir))
        {
            try { Directory.Delete(_tempDir, recursive: true); } catch { }
        }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task McpServer_Initialize_ShouldReturnValidResponse()
    {
        // Arrange
        var server = StartMcpServer();

        // Act - Send MCP initialize request
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
                    name = "ExcelMcp-Test-Client",
                    version = "1.0.0"
                }
            }
        };

        var response = await SendMcpRequestAsync(server, initRequest);

        // Assert
        Assert.NotNull(response);
        var json = JsonDocument.Parse(response);
        Assert.Equal("2.0", json.RootElement.GetProperty("jsonrpc").GetString());
        Assert.Equal(1, json.RootElement.GetProperty("id").GetInt32());

        var result = json.RootElement.GetProperty("result");
        Assert.True(result.TryGetProperty("protocolVersion", out _));
        Assert.True(result.TryGetProperty("serverInfo", out _));
        Assert.True(result.TryGetProperty("capabilities", out _));
    }

    [Fact]
    public async Task McpServer_ListTools_ShouldReturn6ExcelTools()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);

        // Act - Send tools/list request
        var toolsRequest = new
        {
            jsonrpc = "2.0",
            id = 2,
            method = "tools/list",
            @params = new { }
        };

        var response = await SendMcpRequestAsync(server, toolsRequest);

        // Assert
        var json = JsonDocument.Parse(response);
        var tools = json.RootElement.GetProperty("result").GetProperty("tools");

        Assert.Equal(7, tools.GetArrayLength());

        var toolNames = tools.EnumerateArray()
            .Select(t => t.GetProperty("name").GetString())
            .OrderBy(n => n)
            .ToArray();

        Assert.Equal(new[] {
            "excel_cell",
            "excel_connection",
            "excel_file",
            "excel_parameter",
            "excel_powerquery",
            "excel_vba",
            "excel_worksheet"
        }, toolNames);
    }

    [Fact]
    public async Task McpServer_CallExcelFileTool_ShouldCreateFileAndReturnSuccess()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);
        var testFile = Path.Combine(_tempDir, "mcp-test.xlsx");

        // Act - Call excel_file tool to create empty file
        var toolCallRequest = new
        {
            jsonrpc = "2.0",
            id = 3,
            method = "tools/call",
            @params = new
            {
                name = "excel_file",
                arguments = new
                {
                    action = "create-empty",
                    excelPath = testFile
                }
            }
        };

        var response = await SendMcpRequestAsync(server, toolCallRequest);

        // Assert
        var json = JsonDocument.Parse(response);
        var result = json.RootElement.GetProperty("result");

        // Should have content array with text content
        Assert.True(result.TryGetProperty("content", out var content));
        var textContent = content.EnumerateArray().First();
        Assert.Equal("text", textContent.GetProperty("type").GetString());

        var textValue = textContent.GetProperty("text").GetString();
        Assert.NotNull(textValue);
        var resultJson = JsonDocument.Parse(textValue);
        Assert.True(resultJson.RootElement.GetProperty("success").GetBoolean());

        // Verify file was actually created
        Assert.True(File.Exists(testFile));
    }

    [Fact]
    public async Task McpServer_CallInvalidTool_ShouldReturnError()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);

        // Act - Call non-existent tool
        var toolCallRequest = new
        {
            jsonrpc = "2.0",
            id = 4,
            method = "tools/call",
            @params = new
            {
                name = "non_existent_tool",
                arguments = new { }
            }
        };

        var response = await SendMcpRequestAsync(server, toolCallRequest);

        // Assert
        var json = JsonDocument.Parse(response);
        Assert.True(json.RootElement.TryGetProperty("error", out _));
    }

    [Fact]
    public async Task McpServer_ExcelWorksheetTool_ShouldListWorksheets()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);
        var testFile = Path.Combine(_tempDir, "worksheet-test.xlsx");

        // First create file
        await CallExcelTool(server, "excel_file", new { action = "create-empty", excelPath = testFile });

        // Act - List worksheets
        var response = await CallExcelTool(server, "excel_worksheet", new { action = "list", excelPath = testFile });

        // Assert
        var resultJson = JsonDocument.Parse(response);
        Assert.True(resultJson.RootElement.GetProperty("Success").GetBoolean());
        Assert.True(resultJson.RootElement.TryGetProperty("Worksheets", out _));
    }

    [Fact]
    public async Task McpServer_PowerQueryWorkflow_ShouldCreateAndReadQuery()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);
        var testFile = Path.Combine(_tempDir, "powerquery-test.xlsx");
        var queryName = "TestQuery";
        var mCodeFile = Path.Combine(_tempDir, "test-query.pq");

        // Create a simple M code query
        var mCode = @"let
    Source = ""Hello from Power Query!"",
    Output = Source
in
    Output";
        await File.WriteAllTextAsync(mCodeFile, mCode);

        // First create Excel file
        await CallExcelTool(server, "excel_file", new { action = "create-empty", excelPath = testFile });

        // Act - Import Power Query
        var importResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "import",
            excelPath = testFile,
            queryName = queryName,
            sourcePath = mCodeFile
        });

        // Assert import succeeded
        var importJson = JsonDocument.Parse(importResponse);
        Assert.True(importJson.RootElement.GetProperty("Success").GetBoolean());

        // Act - Read the Power Query back
        var viewResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "view",
            excelPath = testFile,
            queryName = queryName
        });

        // Assert view succeeded and contains the M code
        var viewJson = JsonDocument.Parse(viewResponse);
        Assert.True(viewJson.RootElement.GetProperty("Success").GetBoolean());
        Assert.True(viewJson.RootElement.TryGetProperty("MCode", out var formulaElement));

        var retrievedMCode = formulaElement.GetString();
        Assert.NotNull(retrievedMCode);
        Assert.Contains("Hello from Power Query!", retrievedMCode);
        Assert.Contains("let", retrievedMCode);

        // Act - List queries to verify it appears in the list
        var listResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "list",
            excelPath = testFile
        });

        // Assert query appears in list
        var listJson = JsonDocument.Parse(listResponse);
        Assert.True(listJson.RootElement.GetProperty("Success").GetBoolean());
        Assert.True(listJson.RootElement.TryGetProperty("Queries", out var queriesElement));

        var queries = queriesElement.EnumerateArray().Select(q => q.GetProperty("Name").GetString()).ToArray();
        Assert.Contains(queryName, queries);

        _output.WriteLine($"Successfully created and read Power Query '{queryName}'");
        _output.WriteLine($"Retrieved M code: {retrievedMCode}");

        // Act - Delete the Power Query to complete the workflow
        var deleteResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "delete",
            excelPath = testFile,
            queryName = queryName
        });

        // Assert delete succeeded
        var deleteJson = JsonDocument.Parse(deleteResponse);
        Assert.True(deleteJson.RootElement.GetProperty("Success").GetBoolean());

        // Verify query is no longer in the list
        var finalListResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "list",
            excelPath = testFile
        });

        var finalListJson = JsonDocument.Parse(finalListResponse);
        Assert.True(finalListJson.RootElement.GetProperty("Success").GetBoolean());

        if (finalListJson.RootElement.TryGetProperty("queries", out var finalQueriesElement))
        {
            var finalQueries = finalQueriesElement.EnumerateArray().Select(q => q.GetProperty("name").GetString()).ToArray();
            Assert.DoesNotContain(queryName, finalQueries);
        }

        _output.WriteLine($"Successfully deleted Power Query '{queryName}' - complete workflow test passed");
    }

    // Helper Methods
    private Process StartMcpServer()
    {
        var serverExePath = Path.Combine(
            Directory.GetCurrentDirectory(),
            "..", "..", "..", "..", "..", "src", "ExcelMcp.McpServer", "bin", "Debug", "net9.0",
            "Sbroenne.ExcelMcp.McpServer.exe"
        );
        serverExePath = Path.GetFullPath(serverExePath);

        if (!File.Exists(serverExePath))
        {
            // Fallback to DLL execution
            serverExePath = Path.Combine(
                Directory.GetCurrentDirectory(),
                "..", "..", "..", "..", "..", "src", "ExcelMcp.McpServer", "bin", "Debug", "net9.0",
                "Sbroenne.ExcelMcp.McpServer.dll"
            );
            serverExePath = Path.GetFullPath(serverExePath);
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = File.Exists(serverExePath) && serverExePath.EndsWith(".exe") ? serverExePath : "dotnet",
            Arguments = File.Exists(serverExePath) && serverExePath.EndsWith(".exe") ? "" : serverExePath,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        var process = Process.Start(startInfo);
        Assert.NotNull(process);

        _serverProcess = process;
        return process;
    }

    private async Task<string> SendMcpRequestAsync(Process server, object request)
    {
        var json = JsonSerializer.Serialize(request);
        _output.WriteLine($"Sending: {json}");

        await server.StandardInput.WriteLineAsync(json);
        await server.StandardInput.FlushAsync();

        var response = await server.StandardOutput.ReadLineAsync();
        _output.WriteLine($"Received: {response ?? "NULL"}");

        Assert.NotNull(response);
        return response;
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
                clientInfo = new { name = "Test", version = "1.0.0" }
            }
        };

        await SendMcpRequestAsync(server, initRequest);

        // Send initialized notification
        var initializedNotification = new
        {
            jsonrpc = "2.0",
            method = "notifications/initialized",
            @params = new { }
        };

        var json = JsonSerializer.Serialize(initializedNotification);
        await server.StandardInput.WriteLineAsync(json);
        await server.StandardInput.FlushAsync();
    }

    private async Task<string> CallExcelTool(Process server, string toolName, object arguments)
    {
        var toolCallRequest = new
        {
            jsonrpc = "2.0",
            id = Environment.TickCount & 0x7FFFFFFF, // Use tick count for test IDs
            method = "tools/call",
            @params = new
            {
                name = toolName,
                arguments
            }
        };

        var response = await SendMcpRequestAsync(server, toolCallRequest);
        var json = JsonDocument.Parse(response);
        var result = json.RootElement.GetProperty("result");
        var content = result.GetProperty("content").EnumerateArray().First();
        var textValue = content.GetProperty("text").GetString();
        return textValue ?? string.Empty;
    }

}
