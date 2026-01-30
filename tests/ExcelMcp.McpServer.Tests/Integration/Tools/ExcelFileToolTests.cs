// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.Models;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Tests for ExcelFileTool action methods.
/// These tests call the tool methods directly without MCP transport.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "File")]
public class ExcelFileToolTests(ITestOutputHelper output)
{
    [Fact]
    public void CreateAndOpen_ProtectedSystemPath_ReturnsJsonError()
    {
        // Arrange - path that reliably fails (Windows directory is protected)
        var protectedPath = @"C:\Windows\HelloWorld.xlsx";

        // Act
        var result = ExcelFileTool.ExcelFile(
            FileAction.CreateAndOpen,
            excelPath: protectedPath,
            sessionId: null,
            save: false,
            showExcel: false,
            timeoutSeconds: 300);

        output.WriteLine($"Result: {result}");

        // Assert - should return JSON error, not crash the server
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.True(json.TryGetProperty("errorMessage", out var errorMsg));
        Assert.Contains("Cannot create", errorMsg.GetString());
        Assert.True(json.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }

    [Fact]
    public void CreateAndOpen_InvalidPath_ReturnsJsonError()
    {
        // Arrange - use a path that will fail (System32, no permission)
        var invalidPath = @"C:\Windows\System32\test.xlsx";

        // Act
        var result = ExcelFileTool.ExcelFile(
            FileAction.CreateAndOpen,
            excelPath: invalidPath,
            sessionId: null,
            save: false,
            showExcel: false,
            timeoutSeconds: 300);

        output.WriteLine($"Result: {result}");

        // Assert - should return JSON error, not crash
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.True(json.TryGetProperty("errorMessage", out var errorMsg));
        Assert.Contains("Cannot create", errorMsg.GetString());
        Assert.True(json.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }

    [Fact]
    public void CreateAndOpen_NullPath_ReturnsJsonError()
    {
        // Act - null path should be caught and returned as JSON error
        var result = ExcelFileTool.ExcelFile(
            FileAction.CreateAndOpen,
            excelPath: null,
            sessionId: null,
            save: false,
            showExcel: false,
            timeoutSeconds: 300);

        output.WriteLine($"Result: {result}");

        // Assert - should return JSON error (ExecuteToolAction wraps exceptions)
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;

        // ExecuteToolAction uses "ok" and "err" for error responses
        Assert.False(json.GetProperty("ok").GetBoolean());
        Assert.True(json.TryGetProperty("err", out var errorMsg));
        Assert.Contains("excelPath is required", errorMsg.GetString());
    }

    [Fact]
    public void CreateAndOpen_ValidPath_ReturnsSuccessWithSessionId()
    {
        // Arrange - use temp directory
        var tempPath = Path.Join(Path.GetTempPath(), $"ExcelFileToolTest_{Guid.NewGuid():N}.xlsx");
        string? sessionId = null;

        try
        {
            // Act
            var result = ExcelFileTool.ExcelFile(
                FileAction.CreateAndOpen,
                excelPath: tempPath,
                sessionId: null,
                save: false,
                showExcel: false,
                timeoutSeconds: 300);

            output.WriteLine($"Result: {result}");

            // Assert
            Assert.NotNull(result);
            var json = JsonDocument.Parse(result).RootElement;
            Assert.True(json.GetProperty("success").GetBoolean());
            Assert.True(File.Exists(tempPath), "File should have been created");
            Assert.True(json.TryGetProperty("sessionId", out var sessionIdElement));
            sessionId = sessionIdElement.GetString();
            Assert.NotNull(sessionId);
        }
        finally
        {
            // Cleanup - close session first
            if (!string.IsNullOrEmpty(sessionId))
            {
                ExcelFileTool.ExcelFile(
                    FileAction.Close,
                    excelPath: null,
                    sessionId: sessionId,
                    save: false,
                    showExcel: false,
                    timeoutSeconds: 300);
            }

            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    [Fact]
    public void Test_NonExistentFile_ReturnsNotFound()
    {
        // Arrange
        var fakePath = @"C:\NonExistent\fake.xlsx";

        // Act
        var result = ExcelFileTool.ExcelFile(
            FileAction.Test,
            excelPath: fakePath,
            sessionId: null,
            save: false,
            showExcel: false,
            timeoutSeconds: 300);

        output.WriteLine($"Result: {result}");

        // Assert
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.False(json.GetProperty("exists").GetBoolean());
    }
}
