using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

// CA2201: These tests fabricate the COMException and COM-mapped OutOfMemoryException
// that Excel raises nondeterministically, to exercise the pure per-query outcome logic.
#pragma warning disable CA2201

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
public sealed class PowerQueryRefreshAllOutcomeTests
{
    private const int EOutOfMemory = unchecked((int)0x8007000E);

    [Fact]
    public void RefreshQueries_OutOfMemoryFromOneQuery_RecordsFailureAndRefreshesRemainingQueries()
    {
        var attempted = new List<string>();
        var outOfMemory = new OutOfMemoryException(
            "Not enough memory resources are available to complete this operation. (0x8007000E (E_OUTOFMEMORY))");
        outOfMemory.HResult = EOutOfMemory;

        var result = PowerQueryCommands.RefreshQueries(
            ["Countries", "WDI_Long", "WDI_Wide"],
            name =>
            {
                attempted.Add(name);
                return name == "WDI_Long" ? throw outOfMemory : true;
            },
            CancellationToken.None);

        Assert.Equal(["Countries", "WDI_Long", "WDI_Wide"], attempted);
        Assert.False(result.Success);
        Assert.Equal(["Countries", "WDI_Wide"], result.RefreshedQueries);
        Assert.Empty(result.SkippedQueries);
        var failure = Assert.Single(result.FailedQueries);
        Assert.Equal("WDI_Long", failure.QueryName);
        Assert.Equal(nameof(OutOfMemoryException), failure.ExceptionType);
        Assert.Equal("0x8007000E", failure.HResult);
        Assert.Contains("E_OUTOFMEMORY", failure.ErrorMessage, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(result.ErrorMessage));
        Assert.Contains("WDI_Long", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("Countries", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("WDI_Wide", result.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void RefreshQueries_QueriesWithNothingToRefresh_AreSkippedAndResultSucceeds()
    {
        var result = PowerQueryCommands.RefreshQueries(
            ["pCountries", "Countries", "Stage"],
            name => name == "Countries",
            CancellationToken.None);

        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(["Countries"], result.RefreshedQueries);
        Assert.Equal(["pCountries", "Stage"], result.SkippedQueries.Select(s => s.QueryName));
        Assert.All(result.SkippedQueries, s => Assert.False(string.IsNullOrWhiteSpace(s.Reason)));
        Assert.Empty(result.FailedQueries);
        Assert.Contains("pCountries", result.Message, StringComparison.Ordinal);
        Assert.Contains("Stage", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RefreshQueries_NoQueries_Succeeds()
    {
        var result = PowerQueryCommands.RefreshQueries([], _ => true, CancellationToken.None);

        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Empty(result.RefreshedQueries);
        Assert.Empty(result.SkippedQueries);
        Assert.Empty(result.FailedQueries);
    }

    [Fact]
    public void RefreshQueries_MixedOutcomes_ReportsEveryQueryWithCategory()
    {
        var result = PowerQueryCommands.RefreshQueries(
            ["Good", "Engine", "Com", "Param", "Cancelled"],
            name => name switch
            {
                "Engine" => throw new COMException(
                    "[Expression.Error] The name 'Missing' wasn't recognized.", unchecked((int)0x800A03EC)),
                "Com" => throw new COMException("Generic failure", unchecked((int)0x80004005)),
                "Param" => false,
                "Cancelled" => throw new OperationFailureException(
                    OperationFailureCategory.Cancelled, "Power Query refresh for 'Cancelled' was cancelled by Excel."),
                _ => true
            },
            CancellationToken.None);

        Assert.False(result.Success);
        Assert.Equal(["Good"], result.RefreshedQueries);
        Assert.Equal(["Param"], result.SkippedQueries.Select(s => s.QueryName));
        Assert.Equal(["Engine", "Com", "Cancelled"], result.FailedQueries.Select(f => f.QueryName));

        var engine = result.FailedQueries[0];
        Assert.Equal("Expression", engine.ErrorCategory);
        Assert.Equal("0x800A03EC", engine.HResult);
        Assert.Contains("wasn't recognized", engine.ErrorMessage, StringComparison.Ordinal);

        var com = result.FailedQueries[1];
        Assert.Null(com.ErrorCategory);
        Assert.Equal(nameof(COMException), com.ExceptionType);
        Assert.Equal("0x80004005", com.HResult);

        var cancelled = result.FailedQueries[2];
        Assert.Equal(nameof(OperationFailureCategory.Cancelled), cancelled.ErrorCategory);
        Assert.Null(cancelled.HResult);
    }

    [Fact]
    public void RefreshQueries_TimeoutException_PropagatesWithoutAttemptingLaterQueries()
    {
        var attempted = new List<string>();

        Assert.Throws<TimeoutException>(() => PowerQueryCommands.RefreshQueries(
            ["First", "Second"],
            name =>
            {
                attempted.Add(name);
                throw new TimeoutException("timed out");
            },
            CancellationToken.None));

        Assert.Equal(["First"], attempted);
    }

    [Fact]
    public void RefreshQueries_OperationCanceled_Propagates()
    {
        Assert.Throws<OperationCanceledException>(() => PowerQueryCommands.RefreshQueries(
            ["First", "Second"],
            _ => throw new OperationCanceledException(),
            CancellationToken.None));
    }

    [Fact]
    public void RefreshQueries_ComFailureAfterTokenCancelled_PropagatesInsteadOfRecording()
    {
        using var cts = new CancellationTokenSource();
        var attempted = new List<string>();

        var exception = Assert.Throws<COMException>(() => PowerQueryCommands.RefreshQueries(
            ["First", "Second"],
            name =>
            {
                attempted.Add(name);
                cts.Cancel();
                throw new COMException("Call was canceled by the message filter.", unchecked((int)0x80010002));
            },
            cts.Token));

        Assert.Equal(unchecked((int)0x80010002), exception.HResult);
        Assert.Equal(["First"], attempted);
    }

    [Fact]
    public void RefreshQueries_TokenCancelledBeforeNextQuery_StopsLoop()
    {
        using var cts = new CancellationTokenSource();
        var attempted = new List<string>();

        Assert.Throws<OperationCanceledException>(() => PowerQueryCommands.RefreshQueries(
            ["First", "Second"],
            name =>
            {
                attempted.Add(name);
                cts.Cancel();
                return true;
            },
            cts.Token));

        Assert.Equal(["First"], attempted);
    }

    [Theory]
    [InlineData(ResiliencePipelines.RPC_E_DISCONNECTED)]
    [InlineData(ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE)]
    [InlineData(ResiliencePipelines.RPC_E_CALL_FAILED)]
    public void RefreshQueries_FatalExcelDisconnect_Propagates(int hresult)
    {
        var attempted = new List<string>();

        var exception = Assert.Throws<COMException>(() => PowerQueryCommands.RefreshQueries(
            ["First", "Second"],
            name =>
            {
                attempted.Add(name);
                throw new COMException("Excel disconnected", hresult);
            },
            CancellationToken.None));

        Assert.Equal(hresult, exception.HResult);
        Assert.Equal(["First"], attempted);
    }

    [Fact]
    public void RefreshQueries_WrappedFatalExcelDisconnect_Propagates()
    {
        var inner = new COMException("Excel disconnected", ResiliencePipelines.RPC_E_DISCONNECTED);

        var exception = Assert.Throws<InvalidOperationException>(() => PowerQueryCommands.RefreshQueries(
            ["First", "Second"],
            _ => throw new InvalidOperationException("wrapped", inner),
            CancellationToken.None));

        Assert.Same(inner, exception.InnerException);
    }
}
