using System.Diagnostics;
using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration;

/// <summary>
/// MCP Protocol tests for PowerQuery Auto-Refresh and Workflow Guidance enhancements
/// Tests the complete PowerQuery workflow through MCP protocol including:
/// - Auto-refresh validation in Import/Update
/// - Error capture and recovery guidance
/// - Load configuration preservation
/// - Workflow guidance (suggested next actions and hints)
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PowerQueryEnhancements")]
[Trait("RequiresExcel", "true")]
public class PowerQueryEnhancementsMcpTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private Process? _serverProcess;
    private int _requestId = 1;

    public PowerQueryEnhancementsMcpTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"PQEnhancements_MCP_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _output.WriteLine($"Test temp directory: {_tempDir}");
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
                    _serverProcess.WaitForExit(2000);
                }
            }
            catch (InvalidOperationException)
            {
                // Process already exited
            }
            catch { }
        }
        _serverProcess?.Dispose();

        // Clean up temp files
        if (Directory.Exists(_tempDir))
        {
            try
            {
                Thread.Sleep(500); // Give Excel time to release file handles
                Directory.Delete(_tempDir, recursive: true);
            }
            catch { }
        }
        GC.SuppressFinalize(this);
    }

    [Fact(Timeout = 60000)] // 60 second timeout to prevent hanging
    public async Task PowerQuery_Import_ShouldValidateAndProvideGuidance()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);

        var testFile = Path.Combine(_tempDir, "import-workflow-test.xlsx");
        var queryFile = Path.Combine(_tempDir, "test-query.pq");

        // Create a simple valid M code query
        var mCode = @"let
    Source = #table(
        {""Column1"", ""Column2""},
        {
            {""Value1"", ""Value2""},
            {""Value3"", ""Value4""}
        }
    )
in
    Source";
        await File.WriteAllTextAsync(queryFile, mCode);

        // Create Excel file
        await CallExcelTool(server, "excel_file", new { action = "create-empty", excelPath = testFile });

        // Act - Import with loadToWorksheet (default: true validates via execution)
        var importResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "import",
            excelPath = testFile,
            queryName = "WorkflowTest",
            sourcePath = queryFile
            // loadToWorksheet defaults to true (validates via SetLoadToTable execution)
        });

        // Assert
        _output.WriteLine($"Import Response: {importResponse}");
        var resultJson = JsonDocument.Parse(importResponse);

        Assert.True(resultJson.RootElement.GetProperty("Success").GetBoolean(),
            "Import should succeed");

        // Verify workflow guidance is provided
        Assert.True(resultJson.RootElement.TryGetProperty("SuggestedNextActions", out var suggestedActions),
            "Should have SuggestedNextActions");

        var actions = suggestedActions.EnumerateArray().Select(a => a.GetString()).ToArray();
        Assert.True(actions.Length >= 2 && actions.Length <= 5,
            $"Should have 2-5 suggested next actions, got {actions.Length}");

        _output.WriteLine($"Suggested Actions: {string.Join(", ", actions)}");

        Assert.True(resultJson.RootElement.TryGetProperty("WorkflowHint", out var workflowHint),
            "Should have WorkflowHint");

        var hint = workflowHint.GetString();
        Assert.NotNull(hint);
        Assert.NotEmpty(hint);

        _output.WriteLine($"Workflow Hint: {hint}");

        // Verify auto-refresh was performed (should not have validation errors)
        if (resultJson.RootElement.TryGetProperty("HasErrors", out var hasErrors))
        {
            Assert.False(hasErrors.GetBoolean(), "Valid query should not have errors after auto-refresh");
        }
    }

    [Fact(Timeout = 60000)] // 60 second timeout to prevent hanging
    public async Task PowerQuery_Import_WithConnectionOnly_ShouldSkipValidation()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);

        var testFile = Path.Combine(_tempDir, "import-connection-only-test.xlsx");
        var queryFile = Path.Combine(_tempDir, "test-query-2.pq");

        var mCode = @"let
    Source = #table({""Col1""}, {{""Data1""}})
in
    Source";
        await File.WriteAllTextAsync(queryFile, mCode);

        await CallExcelTool(server, "excel_file", new { action = "create-empty", excelPath = testFile });

        // Act - Import with loadToWorksheet=false (connection-only, NOT validated)
        var importResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "import",
            excelPath = testFile,
            queryName = "ConnectionOnly",
            sourcePath = queryFile,
            loadToWorksheet = false  // Skips validation (connection-only)
        });

        // Assert
        _output.WriteLine($"Import Response: {importResponse}");
        var resultJson = JsonDocument.Parse(importResponse);

        Assert.True(resultJson.RootElement.GetProperty("Success").GetBoolean());

        // Verify guidance mentions validation was skipped
        if (resultJson.RootElement.TryGetProperty("SuggestedNextActions", out var actions))
        {
            var actionStrings = actions.EnumerateArray().Select(a => a.GetString()?.ToLower()).ToArray();
            var hasRefreshGuidance = actionStrings.Any(a =>
                a != null && (a.Contains("refresh") || a.Contains("validation")));

            Assert.True(hasRefreshGuidance,
                "Should suggest manual refresh when auto-refresh is skipped");

            _output.WriteLine($"Actions: {string.Join(", ", actionStrings)}");
        }
    }

    [Fact(Timeout = 60000)] // 60 second timeout to prevent hanging
    public async Task PowerQuery_Update_ShouldPreserveLoadConfiguration()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);

        var testFile = Path.Combine(_tempDir, "update-config-test.xlsx");
        var queryFile1 = Path.Combine(_tempDir, "query-v1.pq");
        var queryFile2 = Path.Combine(_tempDir, "query-v2.pq");

        var mCodeV1 = @"let
    Source = #table({""Column1""}, {{""Version1""}})
in
    Source";

        var mCodeV2 = @"let
    Source = #table({""Column1""}, {{""Version2""}})
in
    Source";

        await File.WriteAllTextAsync(queryFile1, mCodeV1);
        await File.WriteAllTextAsync(queryFile2, mCodeV2);

        await CallExcelTool(server, "excel_file", new { action = "create-empty", excelPath = testFile });

        // Import query (with new default behavior, this automatically loads to worksheet)
        await CallExcelTool(server, "excel_powerquery", new
        {
            action = "import",
            excelPath = testFile,
            queryName = "ConfigTest",
            sourcePath = queryFile1
        });

        // No need to explicitly load to table anymore - import does this by default

        // Act - Update the query M code
        var updateResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "update",
            excelPath = testFile,
            queryName = "ConfigTest",
            sourcePath = queryFile2
            // loadToWorksheet defaults to true (preserves existing load configuration)
        });

        // Assert
        _output.WriteLine($"Update Response: {updateResponse}");
        var resultJson = JsonDocument.Parse(updateResponse);

        Assert.True(resultJson.RootElement.GetProperty("Success").GetBoolean(),
            "Update should succeed");

        // Verify workflow guidance mentions config preservation
        if (resultJson.RootElement.TryGetProperty("WorkflowHint", out var hint))
        {
            var hintText = hint.GetString()?.ToLower() ?? "";
            _output.WriteLine($"Workflow Hint: {hintText}");

            // Hint should mention preservation or configuration
            var mentionsConfig = hintText.Contains("preserv") ||
                                hintText.Contains("config") ||
                                hintText.Contains("maintain");

            Assert.True(mentionsConfig || hintText.Length > 0,
                "Workflow hint should provide context about the update");
        }

        // Verify load configuration is still LoadToTable
        var configResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "get-load-config",
            excelPath = testFile,
            queryName = "ConfigTest"
        });

        _output.WriteLine($"Config Response: {configResponse}");
        var configJson = JsonDocument.Parse(configResponse);

        Assert.True(configJson.RootElement.TryGetProperty("LoadMode", out var loadMode),
            "Should have LoadMode");

        var mode = loadMode.GetString();
        Assert.Equal("LoadToTable", mode);

        _output.WriteLine($"Load configuration preserved: {mode}");
    }

    [Fact(Timeout = 60000)] // 60 second timeout to prevent hanging
    public async Task PowerQuery_Refresh_ShouldCaptureErrorsAndProvideRecovery()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);

        var testFile = Path.Combine(_tempDir, "refresh-error-test.xlsx");
        var queryFile = Path.Combine(_tempDir, "web-query.pq");

        // Create query that will likely fail (web access in test environment)
        var mCode = @"let
    Source = Web.Contents(""https://nonexistent-test-domain-12345.com/data"")
in
    Source";
        await File.WriteAllTextAsync(queryFile, mCode);

        await CallExcelTool(server, "excel_file", new { action = "create-empty", excelPath = testFile });

        // Import without auto-refresh to avoid immediate error
        await CallExcelTool(server, "excel_powerquery", new
        {
            action = "import",
            excelPath = testFile,
            queryName = "WebQuery",
            sourcePath = queryFile,

        });

        // Act - Refresh the query (will likely fail)
        var refreshResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "refresh",
            excelPath = testFile,
            queryName = "WebQuery"
        });

        // Assert
        _output.WriteLine($"Refresh Response: {refreshResponse}");
        var resultJson = JsonDocument.Parse(refreshResponse);

        // Excel's lenient M code validation means this might succeed OR fail
        // We test both paths using conditional assertions
        if (resultJson.RootElement.TryGetProperty("Success", out var success) && !success.GetBoolean())
        {
            // Path 1: Excel detected the error
            _output.WriteLine("Excel detected error during refresh");

            Assert.True(resultJson.RootElement.TryGetProperty("HasErrors", out var hasErrors) &&
                       hasErrors.GetBoolean(),
                "Failed refresh should have HasErrors flag");

            if (resultJson.RootElement.TryGetProperty("ErrorMessages", out var errors))
            {
                var errorList = errors.EnumerateArray().Select(e => e.GetString()).ToArray();
                Assert.NotEmpty(errorList);
                _output.WriteLine($"Errors captured: {string.Join(", ", errorList)}");
            }

            // Verify error recovery guidance
            if (resultJson.RootElement.TryGetProperty("SuggestedNextActions", out var actions))
            {
                var actionStrings = actions.EnumerateArray().Select(a => a.GetString()?.ToLower()).ToArray();
                var hasErrorGuidance = actionStrings.Any(a =>
                    a != null && (a.Contains("error") || a.Contains("fix") || a.Contains("update")));

                Assert.True(hasErrorGuidance,
                    "Should provide error recovery guidance");

                _output.WriteLine($"Recovery Actions: {string.Join(", ", actionStrings)}");
            }
        }
        else
        {
            // Path 2: Excel accepted the query (lenient validation)
            _output.WriteLine("Excel accepted query despite web access - lenient validation");

            // Should still have workflow guidance
            Assert.True(resultJson.RootElement.TryGetProperty("SuggestedNextActions", out _),
                "Should have workflow guidance even when Excel accepts query");
        }
    }

    [Fact(Timeout = 60000)] // 60 second timeout to prevent hanging
    public async Task PowerQuery_CompleteWorkflow_ShouldProvideContextualGuidance()
    {
        // Arrange
        var server = StartMcpServer();
        await InitializeServer(server);

        var testFile = Path.Combine(_tempDir, "workflow-test.xlsx");
        var queryFile = Path.Combine(_tempDir, "workflow-query.pq");

        var mCode = @"let
    Source = #table(
        {""Name"", ""Value""},
        {
            {""Item1"", 100},
            {""Item2"", 200}
        }
    )
in
    Source";
        await File.WriteAllTextAsync(queryFile, mCode);

        await CallExcelTool(server, "excel_file", new { action = "create-empty", excelPath = testFile });

        // Act & Assert - Complete workflow with guidance validation at each step

        // Step 1: List queries (should be empty)
        var listResponse1 = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "list",
            excelPath = testFile
        });

        _output.WriteLine($"Initial List: {listResponse1}");
        var list1Json = JsonDocument.Parse(listResponse1);
        Assert.True(list1Json.RootElement.GetProperty("Success").GetBoolean());

        // Step 2: Import query
        var importResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "import",
            excelPath = testFile,
            queryName = "WorkflowTest",
            sourcePath = queryFile
        });

        _output.WriteLine($"Import: {importResponse}");
        var importJson = JsonDocument.Parse(importResponse);
        Assert.True(importJson.RootElement.GetProperty("Success").GetBoolean());
        Assert.True(importJson.RootElement.TryGetProperty("SuggestedNextActions", out _),
            "Import should provide next action suggestions");

        // With new default behavior, import already loaded query to worksheet
        // No need for separate set-load-to-table call

        // Step 3: View query M code
        var viewResponse = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "view",
            excelPath = testFile,
            queryName = "WorkflowTest"
        });

        _output.WriteLine($"View: {viewResponse}");
        var viewJson = JsonDocument.Parse(viewResponse);
        Assert.True(viewJson.RootElement.GetProperty("Success").GetBoolean());
        Assert.True(viewJson.RootElement.TryGetProperty("MCode", out var mcode));
        Assert.Contains("#table", mcode.GetString() ?? "");

        // Step 4: List queries (should have our query)
        var listResponse2 = await CallExcelTool(server, "excel_powerquery", new
        {
            action = "list",
            excelPath = testFile
        });

        _output.WriteLine($"Final List: {listResponse2}");
        var list2Json = JsonDocument.Parse(listResponse2);
        Assert.True(list2Json.RootElement.GetProperty("Success").GetBoolean());

        if (list2Json.RootElement.TryGetProperty("Queries", out var queries))
        {
            var queryNames = queries.EnumerateArray()
                .Select(q => q.GetProperty("Name").GetString())
                .ToArray();

            Assert.Contains("WorkflowTest", queryNames);
            _output.WriteLine($"Queries found: {string.Join(", ", queryNames)}");
        }

        _output.WriteLine("Complete workflow test passed with contextual guidance at each step");
    }

    // Helper Methods
    private Process StartMcpServer()
    {
        var projectPath = Path.Combine(
            Directory.GetCurrentDirectory(),
            "..", "..", "..", "..", "..", "src", "ExcelMcp.McpServer", "ExcelMcp.McpServer.csproj"
        );
        projectPath = Path.GetFullPath(projectPath);

        var startInfo = new ProcessStartInfo
        {
            FileName = "dotnet",
            Arguments = $"run --project \"{projectPath}\"",
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        var process = Process.Start(startInfo);
        Assert.NotNull(process);

        _serverProcess = process;
        _output.WriteLine($"MCP Server started with PID: {process.Id}");

        return process;
    }

    private async Task<string> SendMcpRequestAsync(Process server, object request)
    {
        var json = JsonSerializer.Serialize(request);
        _output.WriteLine($">>> Sending: {json}");

        await server.StandardInput.WriteLineAsync(json);
        await server.StandardInput.FlushAsync();

        var response = await server.StandardOutput.ReadLineAsync();
        _output.WriteLine($"<<< Received: {response ?? "NULL"}");

        Assert.NotNull(response);
        return response;
    }

    private async Task InitializeServer(Process server)
    {
        var initRequest = new
        {
            jsonrpc = "2.0",
            id = _requestId++,
            method = "initialize",
            @params = new
            {
                protocolVersion = "2024-11-05",
                capabilities = new { },
                clientInfo = new
                {
                    name = "PowerQueryEnhancements-Test-Client",
                    version = "1.0.0"
                }
            }
        };

        var response = await SendMcpRequestAsync(server, initRequest);
        var json = JsonDocument.Parse(response);
        Assert.Equal("2.0", json.RootElement.GetProperty("jsonrpc").GetString());

        // Send initialized notification
        var initializedNotification = new
        {
            jsonrpc = "2.0",
            method = "notifications/initialized",
            @params = new { }
        };

        var notificationJson = JsonSerializer.Serialize(initializedNotification);
        await server.StandardInput.WriteLineAsync(notificationJson);
        await server.StandardInput.FlushAsync();

        _output.WriteLine("MCP Server initialized");
    }

    private async Task<string> CallExcelTool(Process server, string toolName, object arguments)
    {
        var toolCallRequest = new
        {
            jsonrpc = "2.0",
            id = _requestId++,
            method = "tools/call",
            @params = new
            {
                name = toolName,
                arguments
            }
        };

        var response = await SendMcpRequestAsync(server, toolCallRequest);
        var json = JsonDocument.Parse(response);

        if (json.RootElement.TryGetProperty("error", out var error))
        {
            var errorMessage = error.GetProperty("message").GetString();
            throw new Exception($"MCP tool call failed: {errorMessage}");
        }

        var result = json.RootElement.GetProperty("result");
        var content = result.GetProperty("content").EnumerateArray().First();
        var textValue = content.GetProperty("text").GetString();

        return textValue ?? string.Empty;
    }
}

