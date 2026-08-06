// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Sbroenne.ExcelMcp.Generated;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Xunit.Abstractions;
using Bridge = Sbroenne.ExcelMcp.McpServer.ServiceBridge.ServiceBridge;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Tests for ExcelFileTool action methods.
/// These tests call the tool methods directly without MCP transport.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "true")]
[Collection("ProgramTransport")]
public class ExcelFileToolTests(ITestOutputHelper output)
{
    private static readonly byte[] Ole2Signature =
    [
        0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1
    ];

    [Fact]
    public void Create_MissingDirectory_ReturnsJsonError()
    {
        var missingDirectory = Path.Join(Path.GetTempPath(), $"Missing_{Guid.NewGuid():N}");
        var invalidPath = Path.Join(missingDirectory, "test.xlsx");

        var result = ExcelFileTool.ExcelFile(
            FileAction.Create,
            path: invalidPath,
            session_id: null,
            save: false,
            show: false,
            timeout_seconds: 300);

        output.WriteLine($"Result: {result}");

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.True(json.TryGetProperty("errorMessage", out var errorMsg));
        Assert.Contains("Directory does not exist", errorMsg.GetString());
        Assert.True(json.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }

    [Fact]
    public void Create_RelativePath_ReturnsJsonError()
    {
        const string invalidPath = @"relative\test.xlsx";

        var result = ExcelFileTool.ExcelFile(
            FileAction.Create,
            path: invalidPath,
            session_id: null,
            save: false,
            show: false,
            timeout_seconds: 300);

        output.WriteLine($"Result: {result}");

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.True(json.TryGetProperty("errorMessage", out var errorMsg));
        Assert.Contains("not an absolute Windows path", errorMsg.GetString());
        Assert.True(json.TryGetProperty("isError", out var isError));
        Assert.True(isError.GetBoolean());
    }

    [Fact]
    public void Create_NullPath_ReturnsJsonError()
    {
        // Act - null path should be caught and returned as JSON error
        var result = ExcelFileTool.ExcelFile(
            FileAction.Create,
            path: null,
            session_id: null,
            save: false,
            show: false,
            timeout_seconds: 300);

        output.WriteLine($"Result: {result}");

        // Assert - should return JSON error (ExecuteToolAction wraps exceptions)
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;

        // ExecuteToolAction uses "success" and "errorMessage" for error responses
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.True(json.TryGetProperty("errorMessage", out var errorMsg));
        Assert.Contains("path is required", errorMsg.GetString());
    }

    [Fact]
    public void Create_ValidPath_ReturnsSuccessWithSessionId()
    {
        // Arrange - use temp directory
        var tempPath = Path.Join(Path.GetTempPath(), $"ExcelFileToolTest_{Guid.NewGuid():N}.xlsx");
        string? sessionId = null;

        try
        {
            // Act
            var result = ExcelFileTool.ExcelFile(
                FileAction.Create,
                path: tempPath,
                session_id: null,
                save: false,
                show: false,
                timeout_seconds: 300);

            output.WriteLine($"Result: {result}");

            // Assert
            Assert.NotNull(result);
            var json = JsonDocument.Parse(result).RootElement;
            Assert.True(json.GetProperty("success").GetBoolean());
            Assert.True(File.Exists(tempPath), "File should have been created");
            Assert.True(json.TryGetProperty("session_id", out var sessionIdElement));
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
                    path: null,
                    session_id: sessionId,
                    save: false,
                    show: false,
                    timeout_seconds: 300);
            }

            if (File.Exists(tempPath))
            {
                try
                {
                    for (int i = 0; i < 10; i++)
                    {
                        try
                        {
                            File.Delete(tempPath);
                            break;
                        }
                        catch (IOException) when (i < 9)
                        {
                            Thread.Sleep(500);
                        }
                        catch (UnauthorizedAccessException) when (i < 9)
                        {
                            Thread.Sleep(500);
                        }
                    }
                }
                catch
                {
                    // Best-effort cleanup for a unique temp file created by this test.
                }
            }
        }
    }

    [Fact]
    public void Preflight_OpenSession_ReturnsRealExcelCapabilityContract()
    {
        var tempPath = Path.Join(Path.GetTempPath(), $"ExcelFileToolPreflight_{Guid.NewGuid():N}.xlsx");
        string? sessionId = null;

        try
        {
            var create = ExcelFileTool.ExcelFile(
                FileAction.Create,
                path: tempPath,
                session_id: null,
                save: false,
                show: false,
                timeout_seconds: 300);
            using var createJson = JsonDocument.Parse(create);
            Assert.True(
                createJson.RootElement.GetProperty("success").GetBoolean(),
                createJson.RootElement.TryGetProperty("errorMessage", out var createError)
                    ? createError.GetString()
                    : create);
            sessionId = createJson.RootElement.GetProperty("session_id").GetString();
            Assert.False(string.IsNullOrWhiteSpace(sessionId));

            var result = ExcelFileTool.ExcelFile(
                FileAction.Preflight,
                path: null,
                session_id: sessionId,
                save: false,
                show: false,
                timeout_seconds: 300);
            using var json = JsonDocument.Parse(result);

            Assert.True(json.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(sessionId, json.RootElement.GetProperty("sessionId").GetString());
            Assert.Equal(tempPath, json.RootElement.GetProperty("filePath").GetString());
            Assert.False(string.IsNullOrWhiteSpace(json.RootElement.GetProperty("excel").GetProperty("version").GetString()));
            Assert.True(json.RootElement.GetProperty("excel").GetProperty("build").GetInt32() > 0);
            Assert.False(string.IsNullOrWhiteSpace(json.RootElement.GetProperty("excel").GetProperty("bitness").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(json.RootElement.GetProperty("excel").GetProperty("operatingSystem").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(json.RootElement.GetProperty("excel").GetProperty("uiLocale").GetString()));
            Assert.True(new[] { "supported", "unsupported" }.Contains(
                json.RootElement.GetProperty("capabilities").GetProperty("formula2").GetProperty("status").GetString()));
            Assert.Equal("notDetermined", json.RootElement.GetProperty("capabilities").GetProperty("pythonInExcel").GetProperty("status").GetString());
            Assert.True(new[] { "supported", "blocked" }.Contains(
                json.RootElement.GetProperty("capabilities").GetProperty("vbaTrust").GetProperty("status").GetString()));
            Assert.True(new[] { "supported", "unsupported", "unavailable" }.Contains(
                json.RootElement.GetProperty("capabilities").GetProperty("powerPivot").GetProperty("status").GetString()));
            Assert.Equal(JsonValueKind.False, json.RootElement.GetProperty("workbook").GetProperty("readOnly").ValueKind);
            Assert.Equal(JsonValueKind.Array, json.RootElement.GetProperty("constraints").ValueKind);
            Assert.NotEqual(DateTime.MinValue, json.RootElement.GetProperty("collectedAtUtc").GetDateTime());
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                ExcelFileTool.ExcelFile(
                    FileAction.Close,
                    path: null,
                    session_id: sessionId,
                    save: false,
                    show: false,
                    timeout_seconds: 300);
            }

            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    [Fact]
    public void SafetyReview_RangeMutation_TravelsThroughGeneratedToolAndRealExcel()
    {
        var tempRoot = Path.Join(Path.GetTempPath(), $"ExcelFileToolSafety_{Guid.NewGuid():N}");
        var workbookPath = Path.Join(tempRoot, "review-forwarding.xlsx");
        var stateRoot = Path.Join(tempRoot, "safety-state");
        string? sessionId = null;

        Directory.CreateDirectory(tempRoot);
        Bridge.SetServiceFactoryForTests(
            () => new ExcelMcpServiceBackend(new ExcelMcpService(stateRoot)));

        try
        {
            using (var create = JsonDocument.Parse(ExcelFileTool.ExcelFile(
                       FileAction.Create,
                       workbookPath,
                       session_id: null,
                       save: false,
                       show: false,
                       timeout_seconds: 300)))
            {
                Assert.True(create.RootElement.GetProperty("success").GetBoolean());
                sessionId = create.RootElement.GetProperty("session_id").GetString();
                Assert.False(string.IsNullOrWhiteSpace(sessionId));
            }

            using (var configure = JsonDocument.Parse(ExcelFileTool.ExcelFile(
                       FileAction.ConfigureSafety,
                       path: null,
                       session_id: sessionId,
                       save: false,
                       show: false,
                       timeout_seconds: 300,
                       review_mode: SafetyReviewMode.Required,
                       checkpoint_mode: SafetyCheckpointMode.OnRequest,
                       journal_mode: SafetyJournalMode.On,
                       verification_mode: SafetyVerificationMode.On,
                       abnormal_shutdown_policy: SafetyAbnormalShutdownPolicy.DiscardWithRecoveryEvidence)))
            {
                Assert.True(configure.RootElement.GetProperty("success").GetBoolean());
            }

            var values = new List<List<object?>> { new() { 42d } };
            using var review = JsonDocument.Parse(ExcelRangeTool.ExcelRange(
                RangeAction.SetValues,
                sessionId!,
                review_only: true,
                checkpoint: true,
                sheet_name: "Sheet1",
                range_address: "A1",
                values: values));
            Assert.False(review.RootElement.GetProperty("executed").GetBoolean());
            var reviewId = review.RootElement.GetProperty("reviewId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(reviewId));

            using var execution = JsonDocument.Parse(ExcelRangeTool.ExcelRange(
                RangeAction.SetValues,
                sessionId!,
                review_id: reviewId,
                checkpoint: true,
                idempotency_key: "generated-safety-retry",
                sheet_name: "Sheet1",
                range_address: "A1",
                values: values));
            Assert.True(execution.RootElement.GetProperty("executed").GetBoolean());
            Assert.Equal(
                "verified",
                execution.RootElement.GetProperty("verification").GetProperty("status").GetString());

            using var readBack = JsonDocument.Parse(ExcelRangeTool.ExcelRange(
                RangeAction.GetValues,
                sessionId!,
                sheet_name: "Sheet1",
                range_address: "A1"));
            Assert.Equal(42d, readBack.RootElement.GetProperty("values")[0][0].GetDouble());
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                _ = ExcelFileTool.ExcelFile(
                    FileAction.Close,
                    path: null,
                    session_id: sessionId,
                    save: false,
                    show: false,
                    timeout_seconds: 300);
            }

            Bridge.ResetForTests();
            if (Directory.Exists(tempRoot))
            {
                Directory.Delete(tempRoot, recursive: true);
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
            path: fakePath,
            session_id: null,
            save: false,
            show: false,
            timeout_seconds: 300);

        output.WriteLine($"Result: {result}");

        // Assert
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result).RootElement;
        Assert.False(json.GetProperty("success").GetBoolean());
        Assert.False(json.GetProperty("exists").GetBoolean());
    }

    [Fact]
    public void Test_IrmSignatureFile_ReturnsIrmMetadata()
    {
        // Arrange
        var tempPath = Path.Join(Path.GetTempPath(), $"ExcelFileTool_Irm_{Guid.NewGuid():N}.xlsx");

        try
        {
            File.WriteAllBytes(tempPath, Ole2Signature);

            // Act
            var result = ExcelFileTool.ExcelFile(
                FileAction.Test,
                path: tempPath,
                session_id: null,
                save: false,
                show: false,
                timeout_seconds: 300);

            output.WriteLine($"Result: {result}");

            // Assert
            var json = JsonDocument.Parse(result).RootElement;
            Assert.True(json.GetProperty("success").GetBoolean());
            Assert.True(json.GetProperty("exists").GetBoolean());
            Assert.True(json.GetProperty("isValid").GetBoolean());
            Assert.True(json.GetProperty("isIrmProtected").GetBoolean());
            Assert.False(json.TryGetProperty("isError", out _));
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



