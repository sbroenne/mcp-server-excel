using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Workflow")]
[Trait("Speed", "Medium")]
public sealed class WorkflowExecutionRealExcelTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;

    public WorkflowExecutionRealExcelTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public async Task ExecutePlan_RunsOrderedOperationsAndReportsTheExactFailure()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            nameof(WorkflowExecutionRealExcelTests),
            nameof(ExecutePlan_RunsOrderedOperationsAndReportsTheExactFailure),
            _fixture.TempDir,
            ".xlsx");
        var stateRoot = Path.Combine(_fixture.TempDir, $"workflow-state-{Guid.NewGuid():N}");

        using var service = new ExcelMcpService(stateRoot);
        var sessionId = await OpenSessionAsync(service, workbookPath);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "workflow.execute-plan",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                operations = new object[]
                {
                    SetValue("A2", 10),
                    SetValue("A3", 20),
                    new
                    {
                        command = "range.set-values",
                        args = new
                        {
                            sheetName = "MissingSheet",
                            rangeAddress = "A1",
                            values = new object?[][] { [30] },
                        },
                    },
                    SetValue("A4", 40),
                },
                stopOnError = true,
            }, ServiceProtocol.JsonOptions),
            Source = "test",
        });

        Assert.False(response.Success);
        using (var result = JsonDocument.Parse(response.Result!))
        {
            Assert.Equal("failed", result.RootElement.GetProperty("outcome").GetString());
            Assert.Equal(2, result.RootElement.GetProperty("failedIndex").GetInt32());
            Assert.Equal(3, result.RootElement.GetProperty("attemptedCount").GetInt32());
            Assert.Equal(4, result.RootElement.GetProperty("steps").GetArrayLength());
            Assert.Equal("notStarted", result.RootElement.GetProperty("steps")[3].GetProperty("status").GetString());
        }

        Assert.Equal([10d, 20d, null], await ReadColumnAsync(service, sessionId, "A2:A4"));
        await CloseSessionAsync(service, sessionId);
    }

    [Fact]
    public async Task ExecutePlan_StopOnErrorFalse_ContinuesAfterKnownFailureAndReportsEveryStep()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            nameof(WorkflowExecutionRealExcelTests),
            "continue-after-failure",
            _fixture.TempDir,
            ".xlsx");
        var stateRoot = Path.Combine(_fixture.TempDir, $"workflow-state-{Guid.NewGuid():N}");

        using var service = new ExcelMcpService(stateRoot);
        var sessionId = await OpenSessionAsync(service, workbookPath);
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "workflow.execute-plan",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                operations = new object[]
                {
                    SetValue("A2", 10),
                    new
                    {
                        command = "range.set-values",
                        args = new { sheetName = "Sheet1", rangeAddress = "A3" },
                    },
                    SetValue("A4", 40),
                },
                stopOnError = false,
            }, ServiceProtocol.JsonOptions),
        });

        Assert.False(response.Success);
        Assert.Equal("PlanFailed", response.ErrorCategory);
        using (var receipt = JsonDocument.Parse(response.Result!))
        {
            Assert.Equal("failed", receipt.RootElement.GetProperty("outcome").GetString());
            Assert.Equal(3, receipt.RootElement.GetProperty("attemptedCount").GetInt32());
            Assert.Equal(2, receipt.RootElement.GetProperty("completedCount").GetInt32());
            Assert.Equal(1, receipt.RootElement.GetProperty("failedIndex").GetInt32());
            Assert.Equal(
                ["completed", "failed", "completed"],
                receipt.RootElement.GetProperty("steps").EnumerateArray()
                    .Select(step => step.GetProperty("status").GetString()!).ToArray());
        }

        Assert.Equal([10d, null, 40d], await ReadColumnAsync(service, sessionId, "A2:A4"));
        await CloseSessionAsync(service, sessionId);
    }

    [Fact]
    public async Task ExecutePlan_CheckpointModeOnce_CreatesOneSharedCheckpointForTwoMutations()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            nameof(WorkflowExecutionRealExcelTests),
            "checkpoint-once",
            _fixture.TempDir,
            ".xlsx");
        var stateRoot = Path.Combine(_fixture.TempDir, $"workflow-state-{Guid.NewGuid():N}");

        using var service = new ExcelMcpService(stateRoot);
        var sessionId = await OpenSessionAsync(service, workbookPath);
        var configure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = "{\"reviewMode\":\"off\",\"checkpointMode\":\"off\",\"journalMode\":\"on\",\"verificationMode\":\"off\",\"abnormalShutdownPolicy\":\"discardWithRecoveryEvidence\"}",
        });
        Assert.True(configure.Success, configure.ErrorMessage);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "workflow.execute-plan",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                operations = new object[]
                {
                    new
                    {
                        command = "range.get-values",
                        args = new { sheetName = "Sheet1", rangeAddress = "A1" },
                    },
                    SetValue("A2", 20),
                    SetValue("A3", 30),
                },
                checkpointMode = "once",
            }, ServiceProtocol.JsonOptions),
        });

        Assert.True(response.Success, response.ErrorMessage);
        using (var receipt = JsonDocument.Parse(response.Result!))
        {
            Assert.Equal("completed", receipt.RootElement.GetProperty("outcome").GetString());
            Assert.Equal(3, receipt.RootElement.GetProperty("attemptedCount").GetInt32());
            Assert.Equal(3, receipt.RootElement.GetProperty("completedCount").GetInt32());
            var checkpoint = receipt.RootElement.GetProperty("checkpoint");
            Assert.False(string.IsNullOrWhiteSpace(checkpoint.GetProperty("recoveryId").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(checkpoint.GetProperty("relativePath").GetString()));
            Assert.Equal(64, checkpoint.GetProperty("sha256").GetString()!.Length);
            Assert.True(checkpoint.GetProperty("size").GetInt64() > 0);
        }

        var journal = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.journal",
            SessionId = sessionId,
        });
        Assert.True(journal.Success, journal.ErrorMessage);
        using (var document = JsonDocument.Parse(journal.Result!))
        {
            var operations = document.RootElement.GetProperty("operations").EnumerateArray().ToArray();
            Assert.Equal(1, operations.Count(operation => operation.TryGetProperty("checkpoint", out var checkpoint) &&
                checkpoint.ValueKind == JsonValueKind.Object && checkpoint.GetProperty("size").GetInt64() > 0));
            Assert.Equal(1, operations.Sum(operation => operation.GetProperty("transitions").EnumerateArray()
                .Count(transition => transition.GetProperty("state").GetString() == "checkpointCreated")));
        }

        Assert.Equal([20d, 30d], await ReadColumnAsync(service, sessionId, "A2:A3"));
        await CloseSessionAsync(service, sessionId);
    }

    [Fact]
    public async Task ExecutePlan_RequiredReview_IsRejectedBeforeMutation()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            nameof(WorkflowExecutionRealExcelTests),
            nameof(ExecutePlan_RequiredReview_IsRejectedBeforeMutation),
            _fixture.TempDir,
            ".xlsx");
        var stateRoot = Path.Combine(_fixture.TempDir, $"workflow-state-{Guid.NewGuid():N}");

        using var service = new ExcelMcpService(stateRoot);
        var sessionId = await OpenSessionAsync(service, workbookPath);
        var configure = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.configure-safety",
            SessionId = sessionId,
            Args = "{\"reviewMode\":\"required\",\"checkpointMode\":\"off\",\"journalMode\":\"on\",\"verificationMode\":\"off\"}",
        });
        Assert.True(configure.Success, configure.ErrorMessage);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "workflow.execute-plan",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { operations = new[] { SetValue("A2", 99) } }, ServiceProtocol.JsonOptions),
        });

        Assert.False(response.Success);
        Assert.Equal("PlanReviewUnavailable", response.ErrorCategory);
        Assert.Equal([null], await ReadColumnAsync(service, sessionId, "A2"));
        await CloseSessionAsync(service, sessionId);
    }

    [Fact]
    public async Task ExecutePlan_PartialInvalidInputWithIdempotencyKey_ReplaysWithoutRepeatingEarlierMutation()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            nameof(WorkflowExecutionRealExcelTests),
            "idempotent-partial",
            _fixture.TempDir,
            ".xlsx");
        var stateRoot = Path.Combine(_fixture.TempDir, $"workflow-state-{Guid.NewGuid():N}");

        using var service = new ExcelMcpService(stateRoot);
        var sessionId = await OpenSessionAsync(service, workbookPath);
        await WriteValuesAsync(service, sessionId, "A1", [["marker"]]);
        var request = new ServiceRequest
        {
            Command = "workflow.execute-plan",
            SessionId = sessionId,
            IdempotencyKey = $"partial-{Guid.NewGuid():N}",
            Args = JsonSerializer.Serialize(new
            {
                operations = new object[]
                {
                    new
                    {
                        command = "rangeedit.insert-rows",
                        args = new { sheetName = "Sheet1", rangeAddress = "1:1" },
                    },
                    new
                    {
                        command = "range.set-values",
                        args = new { sheetName = "Sheet1", rangeAddress = "B1" },
                    },
                },
            }, ServiceProtocol.JsonOptions),
        };

        var first = await service.ProcessAsync(request);
        Assert.False(first.Success);
        Assert.Equal("PlanFailed", first.ErrorCategory);
        await WriteValuesAsync(service, sessionId, "A1", [["sentinel"]]);

        var retry = await service.ProcessAsync(request);
        Assert.Equal(ServiceProtocol.Serialize(first), ServiceProtocol.Serialize(retry));

        var values = await ReadTextColumnAsync(service, sessionId, "A1:A3");
        Assert.Collection(
            values,
            value => Assert.Equal("sentinel", value),
            value => Assert.Equal("marker", value),
            Assert.Null);
        await CloseSessionAsync(service, sessionId);
    }

    [Fact]
    public async Task ExecutePlan_FastCompatiblePlan_UsesOneStaDispatchAndPreservesValues()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            nameof(WorkflowExecutionRealExcelTests),
            "fast-compatible",
            _fixture.TempDir,
            ".xlsx");
        var stateRoot = Path.Combine(_fixture.TempDir, $"workflow-state-{Guid.NewGuid():N}");

        using var service = new ExcelMcpService(stateRoot);
        var sessionId = await OpenSessionAsync(service, workbookPath);
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "workflow.execute-plan",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                operations = new[] { SetValue("A2", 12), SetValue("A3", 13), SetValue("A4", 14) },
                fastMode = true,
            }, ServiceProtocol.JsonOptions),
        });

        Assert.True(response.Success, response.ErrorMessage);
        using (var receipt = JsonDocument.Parse(response.Result!))
        {
            Assert.Equal("completed", receipt.RootElement.GetProperty("outcome").GetString());
            Assert.Equal("fast", receipt.RootElement.GetProperty("executionMode").GetString());
            Assert.True(receipt.RootElement.GetProperty("fastModeRequested").GetBoolean());
            Assert.True(receipt.RootElement.GetProperty("fastModeUsed").GetBoolean());
            Assert.Equal(1, receipt.RootElement.GetProperty("staDispatchCount").GetInt64());
        }

        Assert.Equal([12d, 13d, 14d], await ReadColumnAsync(service, sessionId, "A2:A4"));
        await CloseSessionAsync(service, sessionId);
    }

    [Fact]
    public async Task ExecutePlan_IncompatibleFastPlan_FallsBackBeforeDispatch()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            nameof(WorkflowExecutionRealExcelTests),
            "fast-fallback",
            _fixture.TempDir,
            ".xlsx");
        var stateRoot = Path.Combine(_fixture.TempDir, $"workflow-state-{Guid.NewGuid():N}");

        using var service = new ExcelMcpService(stateRoot);
        var sessionId = await OpenSessionAsync(service, workbookPath);
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "workflow.execute-plan",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                operations = new object[]
                {
                    SetValue("A2", 22),
                    new { command = "calculation.get-mode", args = new { } },
                    SetValue("A3", 23),
                },
                fastMode = true,
            }, ServiceProtocol.JsonOptions),
        });

        Assert.True(response.Success, response.ErrorMessage);
        using (var receipt = JsonDocument.Parse(response.Result!))
        {
            Assert.Equal("sequential-fallback", receipt.RootElement.GetProperty("executionMode").GetString());
            Assert.True(receipt.RootElement.GetProperty("fastModeRequested").GetBoolean());
            Assert.False(receipt.RootElement.GetProperty("fastModeUsed").GetBoolean());
            Assert.Contains("calculation.get-mode", receipt.RootElement.GetProperty("fastModeFallbackReason").GetString());
            Assert.Equal(3, receipt.RootElement.GetProperty("staDispatchCount").GetInt64());
        }

        Assert.Equal([22d, 23d], await ReadColumnAsync(service, sessionId, "A2:A3"));
        await CloseSessionAsync(service, sessionId);
    }

    [Fact]
    public async Task OpenAndDescribe_ReturnsCompactBoundedManifestAndLeavesUsableSession()
    {
        var workbookPath = CoreTestHelper.CreateUniqueTestFile(
            nameof(WorkflowExecutionRealExcelTests),
            nameof(OpenAndDescribe_ReturnsCompactBoundedManifestAndLeavesUsableSession),
            _fixture.TempDir,
            ".xlsx");
        var stateRoot = Path.Combine(_fixture.TempDir, $"workflow-state-{Guid.NewGuid():N}");

        using var service = new ExcelMcpService(stateRoot);
        var seedSessionId = await OpenSessionAsync(service, workbookPath);
        await WriteValuesAsync(service, seedSessionId, "A1:D4",
        [
            ["Name", "Region", "Amount", "Ignored"],
            ["Alpha", "East", 10, "x"],
            ["Beta", "West", 20, "y"],
            ["Gamma", "North", 30, "z"],
        ]);
        await CloseSessionAsync(service, seedSessionId, save: true);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "workflow.open-and-describe",
            Args = JsonSerializer.Serialize(new
            {
                filePath = workbookPath,
                show = false,
                previewRows = 2,
                previewColumns = 3,
            }, ServiceProtocol.JsonOptions),
            Source = "test",
        });

        Assert.True(response.Success, response.ErrorMessage);
        Assert.NotNull(response.Result);
        Assert.True(System.Text.Encoding.UTF8.GetByteCount(response.Result) < 2_500, "Manifest should stay compact.");

        using var manifest = JsonDocument.Parse(response.Result);
        var root = manifest.RootElement;
        var sessionId = root.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        Assert.Equal(Path.GetFullPath(workbookPath), Path.GetFullPath(root.GetProperty("filePath").GetString()!));
        Assert.Equal(2, root.GetProperty("previewRows").GetInt32());
        Assert.Equal(3, root.GetProperty("previewColumns").GetInt32());

        var sheet = Assert.Single(root.GetProperty("sheets").EnumerateArray());
        Assert.Equal("Sheet1", sheet.GetProperty("name").GetString());
        Assert.Equal("$A$1:$D$4", sheet.GetProperty("usedRange").GetString());
        Assert.Equal(4, sheet.GetProperty("rowCount").GetInt32());
        Assert.Equal(4, sheet.GetProperty("columnCount").GetInt32());
        Assert.False(sheet.TryGetProperty("values", out _));

        var preview = sheet.GetProperty("preview").EnumerateArray().ToArray();
        Assert.Equal(2, preview.Length);
        string[] expectedHeaders = ["Name", "Region", "Amount"];
        Assert.Equal(
            expectedHeaders,
            preview[0].EnumerateArray().Select(value => value.GetString()!).ToArray());
        Assert.Equal("Alpha", preview[1][0].GetString());
        Assert.Equal("East", preview[1][1].GetString());
        Assert.Equal(10d, preview[1][2].GetDouble());
        Assert.Equal(1, service.SessionCount);

        await WriteValuesAsync(service, sessionId!, "F1", [["still-open"]]);
        await CloseSessionAsync(service, sessionId!);
        Assert.Equal(0, service.SessionCount);
    }

    private static object SetValue(string address, double value) => new
    {
        command = "range.set-values",
        args = new
        {
            sheetName = "Sheet1",
            rangeAddress = address,
            values = new object?[][] { [value] },
        },
    };

    private static async Task<string> OpenSessionAsync(ExcelMcpService service, string workbookPath)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = workbookPath, show = false }, ServiceProtocol.JsonOptions),
            Source = "test",
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var document = JsonDocument.Parse(response.Result!);
        return document.RootElement.GetProperty("sessionId").GetString()!;
    }

    private static async Task<double?[]> ReadColumnAsync(
        ExcelMcpService service,
        string sessionId,
        string rangeAddress)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName = "Sheet1",
                rangeAddress,
            }, ServiceProtocol.JsonOptions),
            Source = "test",
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var document = JsonDocument.Parse(response.Result!);
        return document.RootElement.GetProperty("values")
            .EnumerateArray()
            .Select(row => row[0].ValueKind == JsonValueKind.Null ? (double?)null : row[0].GetDouble())
            .ToArray();
    }

    private static async Task<string?[]> ReadTextColumnAsync(
        ExcelMcpService service,
        string sessionId,
        string rangeAddress)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName = "Sheet1",
                rangeAddress,
            }, ServiceProtocol.JsonOptions),
            Source = "test",
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var document = JsonDocument.Parse(response.Result!);
        return document.RootElement.GetProperty("values")
            .EnumerateArray()
            .Select(row => row[0].ValueKind == JsonValueKind.Null ? null : row[0].GetString())
            .ToArray();
    }

    private static async Task WriteValuesAsync(
        ExcelMcpService service,
        string sessionId,
        string rangeAddress,
        object?[][] values)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new
            {
                sheetName = "Sheet1",
                rangeAddress,
                values,
            }, ServiceProtocol.JsonOptions),
            Source = "test",
        });
        Assert.True(response.Success, response.ErrorMessage);
    }

    private static async Task CloseSessionAsync(ExcelMcpService service, string sessionId, bool save = false)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { save }, ServiceProtocol.JsonOptions),
            Source = "test",
        });
        Assert.True(response.Success, response.ErrorMessage);
    }
}
