using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "false")]
public sealed class FileToolRecordingContractTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task Test_DispatchesTimeoutToReadOnlyValidationOpen()
    {
        var path = Path.Join(
            Path.GetTempPath(),
            $"file-test-recording-{Guid.NewGuid():N}.xlsx");
        var response = new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                filePath = path,
                exists = true,
                isValid = true,
                canOpen = true
            }, ServiceProtocol.JsonOptions)
        };
        var expectedArgs = JsonSerializer.Serialize(new
        {
            filePath = path,
            timeoutSeconds = 45
        }, ServiceProtocol.JsonOptions);

        var call = await _fixture.CallToolAsync(
            "file_read",
            new Dictionary<string, object?>
            {
                ["action"] = "test",
                ["path"] = path,
                ["timeout_seconds"] = 45
            },
            response,
            "session.test",
            expectedSessionId: null,
            expectedArgsJson: expectedArgs);

        Assert.Null(call.Request.SessionId);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.True(result.RootElement.GetProperty("canOpen").GetBoolean());
    }

    [Theory]
    [InlineData("create", "session.create", false)]
    [InlineData("open", "session.open", true)]
    public async Task CreateAndOpen_DispatchNullSessionAndExplicitDefaults(
        string action,
        string expectedCommand,
        bool createExistingFile)
    {
        var path = Path.Join(
            Path.GetTempPath(),
            $"file-recording-{Guid.NewGuid():N}.xlsx");
        if (createExistingFile)
        {
            await File.WriteAllTextAsync(path, string.Empty);
        }

        try
        {
            var response = new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new
                {
                    success = true,
                    sessionId = "recorded-session",
                    filePath = path
                }, ServiceProtocol.JsonOptions)
            };
            var expectedArgs = JsonSerializer.Serialize(new
            {
                filePath = path,
                show = false,
                timeoutSeconds = 120
            }, ServiceProtocol.JsonOptions);

            var call = await _fixture.CallToolAsync(
                "file",
                new Dictionary<string, object?>
                {
                    ["action"] = action,
                    ["path"] = path,
                    ["show"] = false,
                    ["timeout_seconds"] = 120
                },
                response,
                expectedCommand,
                expectedSessionId: null,
                expectedArgsJson: expectedArgs);

            Assert.Null(call.Request.SessionId);
            using var result = JsonDocument.Parse(call.JsonResult);
            Assert.True(result.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(
                "recorded-session",
                result.RootElement.GetProperty("session_id").GetString());
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public async Task List_DispatchesNullSessionAndNullArguments()
    {
        var response = new ServiceResponse
        {
            Success = true,
            Result = """{"success":true,"sessions":[]}"""
        };

        var call = await _fixture.CallToolAsync(
            "file_read",
            new Dictionary<string, object?>
            {
                ["action"] = "list"
            },
            response,
            "session.list",
            expectedSessionId: null,
            expectedArgsJson: null);

        Assert.Null(call.Request.SessionId);
        Assert.Null(call.Request.Args);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal(0, result.RootElement.GetProperty("sessions").GetArrayLength());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task List_PreservesOperationAndVisibilityFields(bool isExcelVisible)
    {
        var response = new ServiceResponse
        {
            Success = true,
            Result = JsonSerializer.Serialize(new
            {
                success = true,
                count = 1,
                sessions = new[]
                {
                    new
                    {
                        sessionId = "session-list",
                        filePath = @"C:\workbook.xlsx",
                        isExcelVisible,
                        activeOperations = 0,
                        canClose = true
                    }
                }
            }, ServiceProtocol.JsonOptions)
        };

        var call = await _fixture.CallToolAsync(
            "file_read",
            new Dictionary<string, object?>
            {
                ["action"] = "list"
            },
            response,
            "session.list",
            expectedSessionId: null,
            expectedArgsJson: null);

        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal(1, result.RootElement.GetProperty("count").GetInt32());
        var session = Assert.Single(
            result.RootElement.GetProperty("sessions").EnumerateArray());
        Assert.Equal("session-list", session.GetProperty("session_id").GetString());
        Assert.False(session.TryGetProperty("sessionId", out _));
        Assert.Equal(@"C:\workbook.xlsx", session.GetProperty("filePath").GetString());
        Assert.Equal(isExcelVisible, session.GetProperty("isExcelVisible").GetBoolean());
        Assert.Equal(0, session.GetProperty("activeOperations").GetInt32());
        Assert.True(session.GetProperty("canClose").GetBoolean());
    }

    [Fact]
    public async Task Close_DispatchesExactSessionAndSaveDefault()
    {
        var call = await _fixture.CallToolAsync(
            "file",
            new Dictionary<string, object?>
            {
                ["action"] = "close",
                ["session_id"] = "session-close",
                ["save"] = false
            },
            new ServiceResponse
            {
                Success = true,
                Result = """{"success":true,"session_id":"session-close","saved":false}"""
            },
            "session.close",
            "session-close",
            """{"save":false}""");

        Assert.Equal("session-close", call.Request.SessionId);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(
            "session-close",
            result.RootElement.GetProperty("session_id").GetString());
        Assert.False(result.RootElement.GetProperty("saved").GetBoolean());
    }

    [Fact]
    public async Task Close_MissingSessionPreservesStructuredError()
    {
        var call = await _fixture.CallToolAsync(
            "file",
            new Dictionary<string, object?>
            {
                ["action"] = "close",
                ["session_id"] = "missing-session",
                ["save"] = false
            },
            new ServiceResponse
            {
                Success = false,
                Command = "session.close",
                SessionId = "missing-session",
                ErrorCategory = "SessionNotFound",
                ErrorMessage = "Session 'missing-session' not found",
                ExceptionType = "InvalidOperationException"
            },
            "session.close",
            "missing-session",
            """{"save":false}""");

        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.True(result.RootElement.GetProperty("isError").GetBoolean());
        Assert.Equal(
            "SessionNotFound",
            result.RootElement.GetProperty("errorCategory").GetString());
        Assert.Equal(
            "Session 'missing-session' not found",
            result.RootElement.GetProperty("errorMessage").GetString());
        Assert.Equal(
            "InvalidOperationException",
            result.RootElement.GetProperty("exceptionType").GetString());
        Assert.Equal("missing-session", result.RootElement.GetProperty("session_id").GetString());
        Assert.False(result.RootElement.TryGetProperty("sessionId", out _));
        Assert.Equal(result.RootElement.GetRawText(), call.Result.StructuredContent!.Value.GetRawText());
    }

    [Fact]
    public async Task Open_LockedErrorAndRetryPreserveProtocolContract()
    {
        var path = Path.Join(
            Path.GetTempPath(),
            $"locked-recording-{Guid.NewGuid():N}.xlsx");
        await File.WriteAllTextAsync(path, string.Empty);
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["path"] = path,
            ["show"] = false,
            ["timeout_seconds"] = 120
        };
        var expectedArgs = JsonSerializer.Serialize(new
        {
            filePath = path,
            show = false,
            timeoutSeconds = 120
        }, ServiceProtocol.JsonOptions);

        try
        {
            var failedCall = await _fixture.CallToolAsync(
                "file",
                arguments,
                new ServiceResponse
                {
                    Success = false,
                    Command = "session.open",
                    ErrorCategory = "WorkbookAccess",
                    ErrorMessage = "Workbook is already open. Close the file and retry with exclusive access.",
                    ExceptionType = "IOException"
                },
                "session.open",
                expectedSessionId: null,
                expectedArgsJson: expectedArgs);
            using (var failedResult = JsonDocument.Parse(failedCall.JsonResult))
            {
                Assert.False(failedResult.RootElement.GetProperty("success").GetBoolean());
                Assert.True(failedResult.RootElement.GetProperty("isError").GetBoolean());
                Assert.Equal(
                    "WorkbookAccess",
                    failedResult.RootElement.GetProperty("errorCategory").GetString());
                Assert.Equal(
                    "Workbook is already open. Close the file and retry with exclusive access.",
                    failedResult.RootElement.GetProperty("errorMessage").GetString());
                Assert.Equal(
                    "IOException",
                    failedResult.RootElement.GetProperty("exceptionType").GetString());
            }

            var retryCall = await _fixture.CallToolAsync(
                "file",
                arguments,
                new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new
                    {
                        success = true,
                        sessionId = "retry-session",
                        filePath = path
                    }, ServiceProtocol.JsonOptions)
                },
                "session.open",
                expectedSessionId: null,
                expectedArgsJson: expectedArgs);
            using var retryResult = JsonDocument.Parse(retryCall.JsonResult);
            Assert.True(retryResult.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal(
                "retry-session",
                retryResult.RootElement.GetProperty("session_id").GetString());
        }
        finally
        {
            File.Delete(path);
        }
    }
}
