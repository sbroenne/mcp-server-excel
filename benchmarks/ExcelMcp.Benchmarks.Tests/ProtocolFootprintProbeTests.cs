using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Benchmarks.Tests;

[Trait("Layer", "Benchmarks")]
[Trait("Category", "Unit")]
[Trait("Feature", "Benchmarks")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ProtocolFootprintProbeTests
{
    [Fact]
    public void FinalVerificationReceipt_RequiresACompleteBoundedReceipt()
    {
        PromptWorkflowWrite[] writes =
        [
            new("A2", 10d),
            new("A3", 20d),
        ];
        var valid = JsonSerializer.Serialize(new
        {
            outcome = "completed",
            verification = new
            {
                status = "verified",
                sheetName = "Data",
                rangeAddress = "$A$2:$A$3",
                rowCount = 2,
                columnCount = 1,
                cellCount = 2,
                inspectedCellCount = 2,
                nonEmptyCellCount = 2,
                formulaCellCount = 0,
                fingerprint = new string('a', 64),
                preview = new object?[][] { [10d], [20d] },
            },
        });

        ProtocolFootprintProbe.EnsureFinalVerificationReceipt(valid, writes);

        var partial = valid.Replace("\"verified\"", "\"partiallyVerified\"", StringComparison.Ordinal);
        Assert.Throws<InvalidDataException>(() =>
            ProtocolFootprintProbe.EnsureFinalVerificationReceipt(partial, writes));
    }

    [Theory]
    [InlineData("{\"success\":true}", true)]
    [InlineData("{\"success\":false}", false)]
    [InlineData("{\"outcome\":\"completed\",\"completedCount\":8}", true)]
    [InlineData("{\"outcome\":\"failed\",\"completedCount\":7}", false)]
    [InlineData("{\"outcome\":\"unknown\",\"completedCount\":7}", false)]
    [InlineData("{}", false)]
    public void SuccessfulToolResult_AcceptsBooleanEnvelopeOrCompletedCompactPlanReceipt(
        string json,
        bool expected)
    {
        Assert.Equal(expected, ProtocolFootprintProbe.IsSuccessfulToolResult(json));
    }

    [Theory]
    [InlineData("Timeout", null)]
    [InlineData("Cancelled", null)]
    [InlineData("Canceled", null)]
    [InlineData("ExcelProcessDied", null)]
    [InlineData("IdempotencyUnknownOutcome", null)]
    [InlineData("IdempotencyInProgress", null)]
    [InlineData("JournalPersistenceFailed", null)]
    [InlineData("AbortedUnknown", null)]
    [InlineData("SessionInterrupted", null)]
    [InlineData("ServerShutdown", null)]
    [InlineData(null, "The operation outcome is unknown after COM dispatch.")]
    [InlineData(null, "The operation was cancelled after dispatch.")]
    [InlineData(null, "Excel process died before the response was returned.")]
    [InlineData(null, "The Excel process is no longer running.")]
    [InlineData(null, "The connection was disconnected after COM dispatch.")]
    public void ConservativeUnknownOutcome_RecognizesAmbiguousFailures(string? errorCategory, string? message)
    {
        Assert.True(ProtocolFootprintProbe.IsConservativeUnknownOutcome(errorCategory, message));
    }

    [Theory]
    [InlineData("TimeoutBeforeExecution", "The request timed out before execution.")]
    [InlineData("CheckpointFailed", "Checkpoint creation failed; the mutation was not run.")]
    [InlineData("InvalidInput", "The request is invalid.")]
    public void ConservativeUnknownOutcome_DoesNotMisclassifyKnownNonExecution(string? errorCategory, string? message)
    {
        Assert.False(ProtocolFootprintProbe.IsConservativeUnknownOutcome(errorCategory, message));
    }
}
