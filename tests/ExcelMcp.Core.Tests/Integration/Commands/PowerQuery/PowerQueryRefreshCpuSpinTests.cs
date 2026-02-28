using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.PowerQuery;

/// <summary>
/// CPU spin regression tests for Power Query refresh operations.
///
/// These tests measure CPU usage during Power Query refresh and assert it stays
/// below a threshold. The root cause of the CPU spin is in the OleMessageFilter:
/// during blocking COM calls like connection.Refresh() or queryTable.Refresh(false),
/// inbound COM callbacks from MashupHost.exe cause either:
/// - WAITNOPROCESS rejection storm (88% CPU) — v1.8.21 behavior
/// - WAITDEFPROCESS + EnsureScanDefinedEvents spin (97% CPU) — original behavior
///
/// The fix: Smart OleMessageFilter with EnterLongOperation/ExitLongOperation that
/// uses WAITDEFPROCESS + SERVERCALL_RETRYLATER rejection in HandleInComingCall,
/// triggering the caller's RetryRejectedCall backoff mechanism.
///
/// These tests use List.Generate to create non-trivial Power Queries (~10K/50K rows)
/// so the refresh takes long enough to measure CPU accurately.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Slow")]
[Trait("RunType", "OnDemand")]
[Collection("Sequential")] // CPU measurement requires isolation — no parallel tests
public class PowerQueryRefreshCpuSpinTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly ITestOutputHelper _output;

    /// <summary>
    /// CPU threshold as a percentage. The bug shows ~88% CPU (WAITNOPROCESS) or ~97%
    /// (WAITDEFPROCESS). With the fix, CPU should be well under 25%.
    /// Using 25% as a generous threshold to avoid flaky failures from OS scheduling jitter.
    /// </summary>
    private const double CpuThresholdPercent = 25.0;

    /// <summary>
    /// M code that generates ~10,000 rows using List.Generate.
    /// Self-contained (no external data sources), takes several seconds to refresh.
    /// </summary>
    private const string MCode10K = """
        let
            Source = List.Generate(
                () => [i = 0],
                each [i] < 10000,
                each [i = [i] + 1],
                each [
                    ID = [i],
                    Name = "Item_" & Text.From([i]),
                    Value = Number.Round(Number.RandomBetween(1, 10000), 2),
                    Category = if Number.Mod([i], 3) = 0 then "A" else if Number.Mod([i], 3) = 1 then "B" else "C",
                    Date = Date.AddDays(#date(2024, 1, 1), Number.Mod([i], 365))
                ]
            ),
            AsTable = Table.FromRecords(Source)
        in
            AsTable
        """;

    /// <summary>
    /// M code that generates ~50,000 rows for longer refresh duration.
    /// </summary>
    private const string MCode50K = """
        let
            Source = List.Generate(
                () => [i = 0],
                each [i] < 50000,
                each [i = [i] + 1],
                each [
                    ID = [i],
                    Name = "Item_" & Text.From([i]),
                    Value = Number.Round(Number.RandomBetween(1, 10000), 2),
                    Category = if Number.Mod([i], 5) = 0 then "A" else if Number.Mod([i], 5) = 1 then "B" else if Number.Mod([i], 5) = 2 then "C" else if Number.Mod([i], 5) = 3 then "D" else "E",
                    Date = Date.AddDays(#date(2024, 1, 1), Number.Mod([i], 365))
                ]
            ),
            AsTable = Table.FromRecords(Source)
        in
            AsTable
        """;

    public PowerQueryRefreshCpuSpinTests(TempDirectoryFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
        var dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(dataModelCommands);
    }

    /// <summary>
    /// CPU spin regression: Data Model query refresh must not spin the CPU.
    ///
    /// This exercises the connection.Refresh() path (Strategy 2 in RefreshConnectionByQueryName).
    /// Data Model queries can't use QueryTable.Refresh — they go through WorkbookConnection.Refresh
    /// which is the primary path affected by the OleMessageFilter CPU spin.
    ///
    /// Before fix: ~88% CPU (WAITNOPROCESS rejection storm)
    /// After fix: &lt;25% CPU (SERVERCALL_RETRYLATER with proper backoff)
    /// </summary>
    [Fact]
    public void Refresh_DataModelQuery_CpuStaysBelowThreshold()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "CpuSpin_DM_" + Guid.NewGuid().ToString("N")[..8];

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create query loaded to Data Model (connection.Refresh path)
        _powerQueryCommands.Create(batch, queryName, MCode10K, PowerQueryLoadMode.LoadToDataModel);

        // Let initialization settle
        Thread.Sleep(1000);

        // Act — measure CPU during refresh
        var process = Process.GetCurrentProcess();
        var cpuBefore = process.TotalProcessorTime;
        var wallBefore = Stopwatch.GetTimestamp();

        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(5));

        var cpuAfter = process.TotalProcessorTime;
        var wallAfter = Stopwatch.GetTimestamp();

        // Calculate
        var cpuDeltaMs = (cpuAfter - cpuBefore).TotalMilliseconds;
        var wallDeltaMs = Stopwatch.GetElapsedTime(wallBefore, wallAfter).TotalMilliseconds;
        var cpuPercent = (cpuDeltaMs / wallDeltaMs) * 100.0;

        _output.WriteLine($"Query: {queryName} (DataModel, 10K rows)");
        _output.WriteLine($"Wall time: {wallDeltaMs:F0}ms");
        _output.WriteLine($"CPU time: {cpuDeltaMs:F1}ms ({cpuPercent:F1}%)");
        _output.WriteLine($"Refresh success: {result.Success}");

        // Assert — refresh must succeed
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");

        // Assert — CPU must stay below threshold
        Assert.True(cpuPercent < CpuThresholdPercent,
            $"REGRESSION: CPU spin during Data Model refresh! {cpuPercent:F1}% " +
            $"({cpuDeltaMs:F1}ms CPU / {wallDeltaMs:F0}ms wall). " +
            $"Threshold: {CpuThresholdPercent}%. " +
            "The OleMessageFilter is likely not rejecting inbound COM callbacks properly.");
    }

    /// <summary>
    /// CPU spin regression: Worksheet query refresh must not spin the CPU.
    ///
    /// This exercises the queryTable.Refresh(false) path (Strategy 1 in RefreshConnectionByQueryName).
    /// Worksheet queries use QueryTable.Refresh which is also affected by the COM callback storm
    /// from MashupHost during refresh.
    ///
    /// Before fix: elevated CPU from COM callback dispatching
    /// After fix: &lt;25% CPU
    /// </summary>
    [Fact]
    public void Refresh_WorksheetQuery_CpuStaysBelowThreshold()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "CpuSpin_WS_" + Guid.NewGuid().ToString("N")[..8];

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create query loaded to worksheet (queryTable.Refresh path)
        _powerQueryCommands.Create(batch, queryName, MCode10K, PowerQueryLoadMode.LoadToTable);

        Thread.Sleep(1000);

        // Act
        var process = Process.GetCurrentProcess();
        var cpuBefore = process.TotalProcessorTime;
        var wallBefore = Stopwatch.GetTimestamp();

        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(5));

        var cpuAfter = process.TotalProcessorTime;
        var wallAfter = Stopwatch.GetTimestamp();

        var cpuDeltaMs = (cpuAfter - cpuBefore).TotalMilliseconds;
        var wallDeltaMs = Stopwatch.GetElapsedTime(wallBefore, wallAfter).TotalMilliseconds;
        var cpuPercent = (cpuDeltaMs / wallDeltaMs) * 100.0;

        _output.WriteLine($"Query: {queryName} (Worksheet, 10K rows)");
        _output.WriteLine($"Wall time: {wallDeltaMs:F0}ms");
        _output.WriteLine($"CPU time: {cpuDeltaMs:F1}ms ({cpuPercent:F1}%)");
        _output.WriteLine($"Refresh success: {result.Success}");

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");

        Assert.True(cpuPercent < CpuThresholdPercent,
            $"REGRESSION: CPU spin during worksheet refresh! {cpuPercent:F1}% " +
            $"({cpuDeltaMs:F1}ms CPU / {wallDeltaMs:F0}ms wall). " +
            $"Threshold: {CpuThresholdPercent}%.");
    }

    /// <summary>
    /// CPU spin regression: Large Data Model query (50K rows) must not spin the CPU.
    ///
    /// The larger dataset produces a longer refresh duration, giving more time for the
    /// COM callback storm to develop and making the CPU spin more observable.
    /// This is the most sensitive test for the CPU spin bug.
    /// </summary>
    [Fact]
    public void Refresh_DataModelQuery_LargeDataset_CpuStaysBelowThreshold()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        var queryName = "CpuSpin_DM50K_" + Guid.NewGuid().ToString("N")[..8];

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create large query loaded to Data Model
        _powerQueryCommands.Create(batch, queryName, MCode50K, PowerQueryLoadMode.LoadToDataModel);

        Thread.Sleep(1000);

        // Act
        var process = Process.GetCurrentProcess();
        var cpuBefore = process.TotalProcessorTime;
        var wallBefore = Stopwatch.GetTimestamp();

        var result = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(10));

        var cpuAfter = process.TotalProcessorTime;
        var wallAfter = Stopwatch.GetTimestamp();

        var cpuDeltaMs = (cpuAfter - cpuBefore).TotalMilliseconds;
        var wallDeltaMs = Stopwatch.GetElapsedTime(wallBefore, wallAfter).TotalMilliseconds;
        var cpuPercent = (cpuDeltaMs / wallDeltaMs) * 100.0;

        _output.WriteLine($"Query: {queryName} (DataModel, 50K rows)");
        _output.WriteLine($"Wall time: {wallDeltaMs:F0}ms");
        _output.WriteLine($"CPU time: {cpuDeltaMs:F1}ms ({cpuPercent:F1}%)");
        _output.WriteLine($"Refresh success: {result.Success}");

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");

        Assert.True(cpuPercent < CpuThresholdPercent,
            $"REGRESSION: CPU spin during large Data Model refresh! {cpuPercent:F1}% " +
            $"({cpuDeltaMs:F1}ms CPU / {wallDeltaMs:F0}ms wall). " +
            $"Threshold: {CpuThresholdPercent}%. " +
            "50K rows should produce sustained refresh with observable spin pattern.");
    }
}
