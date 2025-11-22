using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.Models;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Integration tests for ExcelConnectionTool MCP server operations.
/// Tests the JSON serialization, parameter handling, and error responses.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public class ExcelConnectionToolTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testExcelFile;
    private readonly string _testCsvFile;

    public ExcelConnectionToolTests(ITestOutputHelper output)
    {
        _output = output;

        // Create temp directory for test files
        _tempDir = Path.Join(Path.GetTempPath(), $"ConnectionToolTest_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Join(_tempDir, "ConnectionTest.xlsx");
        _testCsvFile = Path.Join(_tempDir, "TestData.csv");

        _output.WriteLine($"Test directory: {_tempDir}");
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
        {
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Delete_ExistingTextConnection_ReturnsSuccessJson()
    {
        // Arrange - Create workbook
        _output.WriteLine("Creating test workbook...");
        var createResult = ExcelFileTool.ExcelFile(FileAction.CreateEmpty, _testExcelFile);
        AssertSuccess(createResult, "File creation");

        // Open session
        var openResult = ExcelFileTool.ExcelFile(FileAction.Open, _testExcelFile);
        AssertSuccess(openResult, "Open session");
        var openJson = JsonDocument.Parse(openResult);
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.NotNull(sessionId);
        _output.WriteLine($"Session opened: {sessionId}");

        try
        {
            // Create CSV file
            var csvContent = "Name,Value\nTest,123";
            File.WriteAllText(_testCsvFile, csvContent);

            // Create TEXT connection
            string connectionName = "TestTextConnection";
            string connectionString = $"TEXT;{_testCsvFile}";

            _output.WriteLine($"Creating connection '{connectionName}'...");
            var createConnResult = ExcelConnectionTool.ExcelConnection(
                ConnectionAction.Create,
                _testExcelFile,
                sessionId,
                connectionName: connectionName,
                connectionString: connectionString);
            AssertSuccess(createConnResult, "Create connection");

            // Verify connection exists
            var listBefore = ExcelConnectionTool.ExcelConnection(
                ConnectionAction.List,
                _testExcelFile,
                sessionId);
            AssertSuccess(listBefore, "List connections before delete");
            var listBeforeJson = JsonDocument.Parse(listBefore);
            var connectionsBefore = listBeforeJson.RootElement.GetProperty("Connections").EnumerateArray();
            Assert.Contains(connectionsBefore, c => c.GetProperty("Name").GetString() == connectionName);

            // Act - Delete the connection
            _output.WriteLine($"Deleting connection '{connectionName}'...");
            var deleteResult = ExcelConnectionTool.ExcelConnection(
                ConnectionAction.Delete,
                _testExcelFile,
                sessionId,
                connectionName: connectionName);

            // Assert - Verify JSON response
            AssertSuccess(deleteResult, "Delete connection");
            var deleteJson = JsonDocument.Parse(deleteResult);
            Assert.True(deleteJson.RootElement.GetProperty("Success").GetBoolean());
            _output.WriteLine("Delete operation succeeded");

            // Verify connection no longer exists
            var listAfter = ExcelConnectionTool.ExcelConnection(
                ConnectionAction.List,
                _testExcelFile,
                sessionId);
            AssertSuccess(listAfter, "List connections after delete");
            var listAfterJson = JsonDocument.Parse(listAfter);
            var connectionsAfter = listAfterJson.RootElement.GetProperty("Connections").EnumerateArray();
            Assert.DoesNotContain(connectionsAfter, c => c.GetProperty("Name").GetString() == connectionName);
            _output.WriteLine("Connection successfully removed from list");
        }
        finally
        {
            // Close session
            if (!string.IsNullOrEmpty(sessionId))
            {
                ExcelFileTool.ExcelFile(FileAction.Close, sessionId: sessionId);
            }
        }
    }

    [Fact]
    public void Delete_NonExistentConnection_ReturnsErrorJson()
    {
        // Arrange - Create workbook
        _output.WriteLine("Creating test workbook...");
        var createResult = ExcelFileTool.ExcelFile(FileAction.CreateEmpty, _testExcelFile);
        AssertSuccess(createResult, "File creation");

        // Open session
        var openResult = ExcelFileTool.ExcelFile(FileAction.Open, _testExcelFile);
        AssertSuccess(openResult, "Open session");
        var openJson = JsonDocument.Parse(openResult);
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.NotNull(sessionId);

        try
        {
            // Act - Try to delete non-existent connection
            string connectionName = "NonExistentConnection";
            _output.WriteLine($"Attempting to delete non-existent connection '{connectionName}'...");
            var deleteResult = ExcelConnectionTool.ExcelConnection(
                ConnectionAction.Delete,
                _testExcelFile,
                sessionId,
                connectionName: connectionName);

            // Assert - Should return error JSON with isError flag
            _output.WriteLine($"Delete result: {deleteResult}");
            var deleteJson = JsonDocument.Parse(deleteResult);

            // Check for error response
            Assert.False(deleteJson.RootElement.GetProperty("Success").GetBoolean());
            Assert.True(deleteJson.RootElement.TryGetProperty("ErrorMessage", out var errorMsg));
            var errorMessage = errorMsg.GetString();
            Assert.NotNull(errorMessage);
            Assert.Contains("not found", errorMessage, StringComparison.OrdinalIgnoreCase);
            _output.WriteLine($"Expected error received: {errorMessage}");
        }
        finally
        {
            // Close session
            if (!string.IsNullOrEmpty(sessionId))
            {
                ExcelFileTool.ExcelFile(FileAction.Close, sessionId: sessionId);
            }
        }
    }

    [Fact]
    public void Delete_MissingConnectionName_ReturnsErrorJson()
    {
        // Arrange - Create workbook
        _output.WriteLine("Creating test workbook...");
        var createResult = ExcelFileTool.ExcelFile(FileAction.CreateEmpty, _testExcelFile);
        AssertSuccess(createResult, "File creation");

        // Open session
        var openResult = ExcelFileTool.ExcelFile(FileAction.Open, _testExcelFile);
        AssertSuccess(openResult, "Open session");
        var openJson = JsonDocument.Parse(openResult);
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.NotNull(sessionId);

        try
        {
            // Act - Try to delete with null connection name
            _output.WriteLine("Attempting to delete with null connection name...");

            // Assert - Should throw ArgumentException (parameter validation)
            var exception = Assert.Throws<ArgumentException>(() =>
            {
                ExcelConnectionTool.ExcelConnection(
                    ConnectionAction.Delete,
                    _testExcelFile,
                    sessionId,
                    connectionName: null);
            });

            _output.WriteLine($"Expected exception received: {exception.Message}");
            Assert.Contains("connectionName is required", exception.Message);
        }
        finally
        {
            // Close session
            if (!string.IsNullOrEmpty(sessionId))
            {
                ExcelFileTool.ExcelFile(FileAction.Close, sessionId: sessionId);
            }
        }
    }

    [Fact]
    public void Delete_MultipleConnections_RemovesOnlySpecified()
    {
        // Arrange - Create workbook
        _output.WriteLine("Creating test workbook...");
        var createResult = ExcelFileTool.ExcelFile(FileAction.CreateEmpty, _testExcelFile);
        AssertSuccess(createResult, "File creation");

        // Open session
        var openResult = ExcelFileTool.ExcelFile(FileAction.Open, _testExcelFile);
        AssertSuccess(openResult, "Open session");
        var openJson = JsonDocument.Parse(openResult);
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.NotNull(sessionId);

        try
        {
            // Create multiple CSV files and connections
            var csv1Path = Path.Join(_tempDir, "data1.csv");
            var csv2Path = Path.Join(_tempDir, "data2.csv");
            var csv3Path = Path.Join(_tempDir, "data3.csv");

            File.WriteAllText(csv1Path, "A,B\n1,2");
            File.WriteAllText(csv2Path, "C,D\n3,4");
            File.WriteAllText(csv3Path, "E,F\n5,6");

            string conn1 = "Connection1";
            string conn2 = "Connection2";
            string conn3 = "Connection3";

            _output.WriteLine("Creating three connections...");
            AssertSuccess(ExcelConnectionTool.ExcelConnection(
                ConnectionAction.Create, _testExcelFile, sessionId,
                connectionName: conn1, connectionString: $"TEXT;{csv1Path}"), "Create conn1");

            AssertSuccess(ExcelConnectionTool.ExcelConnection(
                ConnectionAction.Create, _testExcelFile, sessionId,
                connectionName: conn2, connectionString: $"TEXT;{csv2Path}"), "Create conn2");

            AssertSuccess(ExcelConnectionTool.ExcelConnection(
                ConnectionAction.Create, _testExcelFile, sessionId,
                connectionName: conn3, connectionString: $"TEXT;{csv3Path}"), "Create conn3");

            // Act - Delete only the second connection
            _output.WriteLine($"Deleting '{conn2}'...");
            var deleteResult = ExcelConnectionTool.ExcelConnection(
                ConnectionAction.Delete,
                _testExcelFile,
                sessionId,
                connectionName: conn2);
            AssertSuccess(deleteResult, "Delete conn2");

            // Assert - Verify only conn2 is deleted
            var listResult = ExcelConnectionTool.ExcelConnection(
                ConnectionAction.List,
                _testExcelFile,
                sessionId);
            AssertSuccess(listResult, "List after delete");

            var listJson = JsonDocument.Parse(listResult);
            var connections = listJson.RootElement.GetProperty("Connections").EnumerateArray().ToList();

            Assert.Contains(connections, c => c.GetProperty("Name").GetString() == conn1);
            Assert.DoesNotContain(connections, c => c.GetProperty("Name").GetString() == conn2);
            Assert.Contains(connections, c => c.GetProperty("Name").GetString() == conn3);

            _output.WriteLine("Verified: Only specified connection was deleted");
        }
        finally
        {
            // Close session
            if (!string.IsNullOrEmpty(sessionId))
            {
                ExcelFileTool.ExcelFile(FileAction.Close, sessionId: sessionId);
            }
        }
    }

    /// <summary>
    /// Helper method to assert operation success and provide clear error messages.
    /// </summary>
    private void AssertSuccess(string jsonResult, string operationName)
    {
        Assert.NotNull(jsonResult);

        try
        {
            var json = JsonDocument.Parse(jsonResult);

            // Check for Success property (most operations)
            if (json.RootElement.TryGetProperty("Success", out var success))
            {
                if (!success.GetBoolean())
                {
                    var errorMsg = json.RootElement.TryGetProperty("ErrorMessage", out var errProp)
                        ? errProp.GetString()
                        : "Unknown error";
                    Assert.Fail($"{operationName} returned Success=false: {errorMsg}");
                }
            }
            // Check for success property (lowercase - batch operations)
            else if (json.RootElement.TryGetProperty("success", out var successLower))
            {
                if (!successLower.GetBoolean())
                {
                    var errorMsg = json.RootElement.TryGetProperty("errorMessage", out var errProp)
                        ? errProp.GetString()
                        : "Unknown error";
                    Assert.Fail($"{operationName} returned success=false: {errorMsg}");
                }
            }

            _output.WriteLine($"  ✓ {operationName} succeeded");
        }
        catch (JsonException ex)
        {
            Assert.Fail($"{operationName} returned invalid JSON: {ex.Message}\nResponse: {jsonResult}");
        }
    }
}
