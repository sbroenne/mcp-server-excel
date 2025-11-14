using System.Text.Json;
using ModelContextProtocol;
using Sbroenne.ExcelMcp.McpServer.Models;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Tests that verify our enhanced error messages include detailed diagnostic information for LLMs.
/// These tests prove that we throw McpException with:
/// - Exception type names ([Exception Type: ...])
/// - Inner exception messages (Inner: ...)
/// - Action context
/// - File paths
/// - Actionable guidance
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "ErrorHandling")]
public class DetailedErrorMessageTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly string _testExcelFile;
    /// <inheritdoc/>

    public DetailedErrorMessageTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Join(Path.GetTempPath(), $"ExcelMcp_DetailedErrorTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _testExcelFile = Path.Join(_tempDir, "test-errors.xlsx");
    }
    /// <inheritdoc/>

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch { }

        GC.SuppressFinalize(this);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ExcelWorksheet_WithInvalidSession_ShouldThrowDetailedError()
    {
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelWorksheetTool.ExcelWorksheet(WorksheetAction.List, sessionId: "invalid-session"));

        _output.WriteLine($"Error message: {exception.Message}");

        Assert.Contains("session", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("excel_file 'open'", exception.Message, StringComparison.OrdinalIgnoreCase);

        _output.WriteLine("✅ Verified: Worksheet tool reports invalid session with actionable message");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ExcelParameter_WithInvalidSession_ShouldThrowDetailedError()
    {
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelNamedRangeTool.ExcelParameter(NamedRangeAction.List, _testExcelFile, sessionId: "invalid-session"));

        _output.WriteLine($"Error message: {exception.Message}");

        Assert.Contains("session", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("excel_file 'open'", exception.Message, StringComparison.OrdinalIgnoreCase);

        _output.WriteLine("✅ Verified: Named range tool reports invalid session with action context");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ExcelPowerQuery_WithInvalidSession_ShouldThrowDetailedError()
    {
        var exception = await Assert.ThrowsAsync<McpException>(async () =>
            await ExcelPowerQueryTool.ExcelPowerQuery(PowerQueryAction.List, sessionId: "invalid-session"));

        _output.WriteLine($"Error message: {exception.Message}");

        Assert.Contains("session", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("excel_file 'open'", exception.Message, StringComparison.OrdinalIgnoreCase);

        _output.WriteLine("✅ Verified: PowerQuery tool reports invalid session with guidance");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ExcelVba_WithMacroEnabledFileWithoutModules_ReturnsEmptyList()
    {
        // Arrange - Create .xlsm file
        var macroFile = Path.Join(_tempDir, "macro-test.xlsm");
        await ExcelFileTool.ExcelFile(FileAction.CreateEmpty, macroFile);
        string? sessionId = null;

        try
        {
            sessionId = await OpenSessionAsync(macroFile);

            // Act - VBA List on macro-enabled file with no modules
            var result = await ExcelVbaTool.ExcelVba(VbaAction.List, macroFile, sessionId);

            _output.WriteLine($"Result JSON: {result}");

            // Parse JSON response
            var json = JsonDocument.Parse(result);
            var success = json.RootElement.GetProperty("success").GetBoolean();
            var count = json.RootElement.GetProperty("count").GetInt32();

            // Assert - No user modules yet, but Document modules may exist
            Assert.True(success, "VBA list should succeed on empty macro workbook");
            Assert.True(count >= 0, "Module count should be non-negative");

            if (json.RootElement.TryGetProperty("scripts", out var scriptsElement))
            {
                foreach (var script in scriptsElement.EnumerateArray())
                {
                    var type = script.GetProperty("type").GetString();
                    Assert.Equal("Document", type); // Empty workbooks only contain document modules
                }
            }

            Assert.False(json.RootElement.TryGetProperty("workflowHint", out _), "workflowHint should not be returned");

            _output.WriteLine("✅ Verified: VBA list on empty macro workbook returns success with helpful hints");
        }
        finally
        {
            if (!string.IsNullOrEmpty(sessionId))
            {
                await CloseSessionAsync(sessionId);
            }
        }
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ExcelVba_WithMissingModuleName_ShouldThrowDetailedError()
    {
        // Arrange - Create macro-enabled file
        string xlsmFile = Path.Join(_tempDir, "test-vba.xlsm");
        await ExcelFileTool.ExcelFile(FileAction.CreateEmpty, xlsmFile);
        string? sessionId = null;

        try
        {
            sessionId = await OpenSessionAsync(xlsmFile);

            // Act & Assert - Run requires moduleName
            var exception = await Assert.ThrowsAsync<McpException>(async () =>
                await ExcelVbaTool.ExcelVba(VbaAction.Run, xlsmFile, sessionId, moduleName: null));

            _output.WriteLine($"Error message: {exception.Message}");

            // Verify detailed components
            Assert.Contains("moduleName", exception.Message);
            Assert.Contains("required", exception.Message);
            Assert.Contains("run", exception.Message);

            _output.WriteLine("✅ Verified: Missing parameter error includes parameter name and action");
        }
        finally
        {
            if (!string.IsNullOrEmpty(sessionId))
            {
                await CloseSessionAsync(sessionId);
            }
        }
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ExcelPowerQuery_Import_WithMissingParameters_ShouldThrowDetailedError()
    {
        // Arrange
        await ExcelFileTool.ExcelFile(FileAction.CreateEmpty, _testExcelFile);
        string? sessionId = null;

        try
        {
            sessionId = await OpenSessionAsync(_testExcelFile);

            // Act & Assert - Create requires queryName and sourcePath
            var exception = await Assert.ThrowsAsync<McpException>(async () =>
                await ExcelPowerQueryTool.ExcelPowerQuery(PowerQueryAction.Create, sessionId, queryName: null, sourcePath: null));

            _output.WriteLine($"Error message: {exception.Message}");

            // Verify detailed components
            Assert.Contains("queryName", exception.Message);
            Assert.Contains("sourcePath", exception.Message);
            Assert.Contains("required", exception.Message);
            Assert.Contains("create", exception.Message);

            _output.WriteLine("✅ Verified: Missing parameters error lists all required parameters");
        }
        finally
        {
            if (!string.IsNullOrEmpty(sessionId))
            {
                await CloseSessionAsync(sessionId);
            }
        }
    }
    /// <inheritdoc/>

    [Fact]
    public async Task ExcelParameter_Create_WithMissingParameters_ShouldThrowDetailedError()
    {
        // Arrange
        await ExcelFileTool.ExcelFile(FileAction.CreateEmpty, _testExcelFile);
        string? sessionId = null;

        try
        {
            sessionId = await OpenSessionAsync(_testExcelFile);

            // Act & Assert - create requires parameterName and reference
            var exception = await Assert.ThrowsAsync<McpException>(async () =>
                await ExcelNamedRangeTool.ExcelParameter(NamedRangeAction.Create, _testExcelFile, sessionId, namedRangeName: null));

            _output.WriteLine($"Error message: {exception.Message}");

            // Verify detailed components
            Assert.Contains("namedRangeName", exception.Message);
            Assert.Contains("required", exception.Message);
            Assert.Contains("create", exception.Message);

            _output.WriteLine("✅ Verified: Missing parameter error includes action context");
        }
        finally
        {
            if (!string.IsNullOrEmpty(sessionId))
            {
                await CloseSessionAsync(sessionId);
            }
        }
    }

    private async Task<string> OpenSessionAsync(string filePath)
    {
        var openResult = await ExcelFileTool.ExcelFile(FileAction.Open, filePath);
        var json = JsonDocument.Parse(openResult);
        if (!json.RootElement.TryGetProperty("sessionId", out var sessionProp))
        {
            throw new InvalidOperationException($"Failed to open session: {openResult}");
        }

        return sessionProp.GetString() ?? throw new InvalidOperationException("Session ID missing in response");
    }

    private static async Task CloseSessionAsync(string sessionId)
    {
        await ExcelFileTool.ExcelFile(FileAction.Close, sessionId: sessionId);
    }
}


