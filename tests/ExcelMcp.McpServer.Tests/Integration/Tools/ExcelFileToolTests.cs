// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Tests for ExcelFileTool action methods.
/// These tests use production registration and MCP transport.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "File")]
[Collection("ProgramTransport")]
[Trait("RequiresExcel", "false")]
public class ExcelFileToolTests(ITestOutputHelper output) : McpIntegrationTestBase(output, "FileValidationClient")
{
    [Fact]
    public async Task Create_MissingDirectory_ReturnsJsonError()
    {
        var missingDirectory = Path.Join(Path.GetTempPath(), $"Missing_{Guid.NewGuid():N}");
        var invalidPath = Path.Join(missingDirectory, "test.xlsx");

        var result = await CallToolAsync("file", new() { ["action"] = "create", ["file_path"] = invalidPath, ["timeout_seconds"] = 300 });

        Output.WriteLine($"Result: {result}");

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.True(json.TryGetProperty("errorMessage", out var errorMsg));
        Assert.Contains("Directory does not exist", errorMsg.GetString());
        Assert.False(json.TryGetProperty("isError", out _));
    }

    [Fact]
    public async Task Create_RelativePath_ReturnsJsonError()
    {
        const string invalidPath = @"relative\test.xlsx";

        var result = await CallToolAsync("file", new() { ["action"] = "create", ["file_path"] = invalidPath, ["timeout_seconds"] = 300 });

        Output.WriteLine($"Result: {result}");

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.True(json.TryGetProperty("errorMessage", out var errorMsg));
        Assert.Contains("not an absolute Windows path", errorMsg.GetString());
        Assert.False(json.TryGetProperty("isError", out _));
    }

    [Fact]
    public async Task Create_NullPath_ReturnsJsonError()
    {
        // Act - null path should be caught and returned as JSON error
        var result = await CallToolAsync("file", new() { ["action"] = "create", ["file_path"] = null, ["timeout_seconds"] = 300 });

        Output.WriteLine($"Result: {result}");

        // Assert - should return JSON error (ExecuteToolAction wraps exceptions)
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;

        // ExecuteToolAction uses "success" and "errorMessage" for error responses
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.True(json.TryGetProperty("errorMessage", out var errorMsg));
        Assert.Contains("file_path", errorMsg.GetString());
        Assert.Contains("required", errorMsg.GetString());
    }

    [Fact]
    public async Task Test_NonExistentFile_ReturnsNotFound()
    {
        // Arrange
        var fakePath = @"C:\NonExistent\fake.xlsx";

        // Act
        var result = await CallToolAsync("file_read", new() { ["action"] = "test", ["file_path"] = fakePath, ["timeout_seconds"] = 300 });

        Output.WriteLine($"Result: {result}");

        // Assert
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.False(json.GetProperty("exists").GetBoolean());
    }

    [Theory]
    [InlineData("\tDRMDataSpace")]
    [InlineData("DRMEncryptedDataSpace")]
    public async Task Test_IrmDataSpaceFile_ReturnsIrmMetadata(string dataSpaceName)
    {
        // Arrange
        var tempPath = Path.Join(Path.GetTempPath(), $"ExcelFileTool_Irm_{Guid.NewGuid():N}.xlsx");

        try
        {
            Sbroenne.ExcelMcp.Tests.Helpers.OleDataSpaceTestFile.Write(
                tempPath,
                dataSpaceName);

            // Act
            var result = await CallToolAsync("file_read", new() { ["action"] = "test", ["file_path"] = tempPath, ["timeout_seconds"] = 300 });

            Output.WriteLine($"Result: {result}");

            // Assert
            var json = JsonDocument.Parse(result).RootElement;
            Assert.False(json.GetProperty("success").GetBoolean());
            Assert.True(json.GetProperty("exists").GetBoolean());
            Assert.False(json.GetProperty("isValid").GetBoolean());
            Assert.False(json.GetProperty("canOpen").GetBoolean());
            Assert.True(json.GetProperty("isIrmProtected").GetBoolean());
            Assert.False(json.GetProperty("willOpenReadOnly").GetBoolean());
            Assert.True(json.GetProperty("requiresVisibleSession").GetBoolean());
            Assert.False(json.TryGetProperty("isError", out _),
                "IRM preflight is a diagnostic result, not a tool execution failure.");
            Assert.Contains("interactive Excel", json.GetProperty("message").GetString(), StringComparison.Ordinal);
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }
}
