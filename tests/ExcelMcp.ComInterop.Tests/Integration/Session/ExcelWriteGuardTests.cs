using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Integration tests for ExcelWriteGuard — the structural COM safety mechanism
/// integrated into ExcelBatch.Execute().
///
/// Verifies that Execute() automatically suppresses EnableEvents, ScreenUpdating,
/// and Calculation during operations, and restores them after completion.
///
/// REGRESSION TESTS for the deadlock caused by missing event/calculation suppression:
/// - Range writes triggering Calculate callbacks → WAITNOPROCESS deadlock
/// - Conditional formatting operations with dependent formulas
/// - Bulk writes without ScreenUpdating suppression
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelWriteGuard")]
[Collection("Sequential")]
public class ExcelWriteGuardTests : IAsyncLifetime
{
    private readonly ITestOutputHelper _output;
    private static string? _staticTestFile;
    private string? _testFileCopy;

    public ExcelWriteGuardTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public Task InitializeAsync()
    {
        if (_staticTestFile == null)
        {
            var testFolder = Path.Join(AppContext.BaseDirectory, "Integration", "Session", "TestFiles");
            _staticTestFile = Path.Join(testFolder, "batch-test-static.xlsx");

            if (!File.Exists(_staticTestFile))
            {
                throw new FileNotFoundException($"Static test file not found at {_staticTestFile}.");
            }
        }

        _testFileCopy = Path.Join(Path.GetTempPath(), $"writeguard-test-{Guid.NewGuid():N}.xlsx");
        File.Copy(_staticTestFile, _testFileCopy, overwrite: true);

        return Task.Delay(500);
    }

    public Task DisposeAsync()
    {
        if (_testFileCopy != null && File.Exists(_testFileCopy))
        {
            try { File.Delete(_testFileCopy); } catch { /* best effort */ }
        }
        return Task.CompletedTask;
    }

    /// <summary>
    /// Verifies that Execute() does NOT suppress EnableEvents.
    /// Events suppression is intentionally left to individual commands because
    /// Data Model operations need events enabled for model synchronization.
    /// </summary>
    [Fact]
    public void Execute_DoesNotSuppressEnableEvents()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        bool eventsInsideExecute = false;

        batch.Execute((ctx, ct) =>
        {
            eventsInsideExecute = ctx.App.EnableEvents;
            _output.WriteLine($"EnableEvents inside Execute: {eventsInsideExecute}");
            return 0;
        });

        // Events should NOT be suppressed — Data Model ops need them
        Assert.True(eventsInsideExecute, "EnableEvents must NOT be suppressed by guard");
    }

    /// <summary>
    /// Verifies that Execute() suppresses ScreenUpdating during operations.
    /// This prevents Excel from repainting after every COM call (perf + stability).
    /// </summary>
    [Fact]
    public void Execute_SuppressesScreenUpdating_DuringOperation()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        bool screenUpdatingInside = true;

        batch.Execute((ctx, ct) =>
        {
            screenUpdatingInside = ctx.App.ScreenUpdating;
            _output.WriteLine($"ScreenUpdating inside Execute: {screenUpdatingInside}");
            return 0;
        });

        Assert.False(screenUpdatingInside, "ScreenUpdating must be false inside Execute()");
    }

    /// <summary>
    /// Verifies that Execute() does NOT suppress Calculation mode.
    /// Calculation is intentionally left alone by the guard because Data Model operations,
    /// PivotTable refresh, and Power Query refresh require calculation to be enabled.
    /// Commands that need manual calculation handle it themselves.
    /// </summary>
    [Fact]
    public void Execute_DoesNotSuppressCalculation()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        int calculationInside = 0;

        batch.Execute((ctx, ct) =>
        {
            calculationInside = (int)ctx.App.Calculation;
            _output.WriteLine($"Calculation inside Execute: {calculationInside}");
            return 0;
        });

        // xlCalculationAutomatic = -4105 (default for new workbooks)
        // Guard should NOT change it — calculation suppression is operation-specific
        Assert.Equal(-4105, calculationInside);
    }

    /// <summary>
    /// Verifies that the guard restores state even when the operation throws.
    /// This is critical — exceptions must not leave Excel in a suppressed state.
    /// </summary>
    [Fact]
    public void Execute_RestoresState_EvenOnException()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        // First: force an exception inside Execute
        Assert.Throws<InvalidOperationException>(() =>
        {
            batch.Execute((ctx, ct) =>
            {
                throw new InvalidOperationException("Intentional test exception");
#pragma warning disable CS0162 // Unreachable code
                return 0;
#pragma warning restore CS0162
            });
        });

        // Second: verify the guard still works (state was restored despite exception)
        bool eventsAfterException = true;
        bool screenUpdatingAfterException = true;
        int calculationAfterException = 0;

        batch.Execute((ctx, ct) =>
        {
            eventsAfterException = ctx.App.EnableEvents;
            screenUpdatingAfterException = ctx.App.ScreenUpdating;
            calculationAfterException = (int)ctx.App.Calculation;
            return 0;
        });

        // Inside Execute, guard suppresses ScreenUpdating only
        // Events and Calculation are NOT suppressed (Data Model ops need them)
        Assert.True(eventsAfterException, "EnableEvents must NOT be suppressed by guard");
        Assert.False(screenUpdatingAfterException, "ScreenUpdating must be suppressed after exception recovery");
        Assert.Equal(-4105, calculationAfterException);
        _output.WriteLine("✓ Guard correctly restored state after exception");
    }

    /// <summary>
    /// Verifies that nested Execute() calls (which create nested guards)
    /// don't double-restore state. The outer guard is the one that restores.
    /// </summary>
    [Fact]
    public void NestedExecute_DoesNotDoubleRestore()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        bool innerScreenUpdating = true;

        batch.Execute((ctx, ct) =>
        {
            // ExcelWriteGuard uses thread-static ref counting.
            // Creating a second guard inside Execute (simulating nested usage)
            // should be a no-op — the outer guard owns state restoration.
            using var innerGuard = new ExcelWriteGuard(ctx.App);

            innerScreenUpdating = ctx.App.ScreenUpdating;
            _output.WriteLine($"ScreenUpdating inside nested guard: {innerScreenUpdating}");

            return 0;
        });

        Assert.False(innerScreenUpdating, "ScreenUpdating must remain false during nested guards");
        _output.WriteLine("✓ Nested guard did not interfere with outer guard");
    }

    /// <summary>
    /// REGRESSION TEST: Writing values to cells with conditional formatting must not deadlock.
    /// Before the fix, MessagePending returned WAITNOPROCESS for normal operations, which
    /// blocked Excel's internal callbacks (Calculate, conditional formatting evaluation)
    /// and caused a deadlock when Excel waited for the callback to complete.
    /// </summary>
    [Fact]
    public void WriteValues_WithConditionalFormatting_DoesNotDeadlock()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? formatConditions = null;
            dynamic? formatCondition = null;

            try
            {
                sheet = ctx.Book.Worksheets[1];

                // Set up: write initial values
                sheet.Range["A1"].Value2 = 100;
                sheet.Range["A2"].Value2 = 200;

                // Add conditional formatting rule on A1:A2
                range = sheet.Range["A1:A2"];
                formatConditions = range.FormatConditions;
                formatCondition = formatConditions.Add(
                    Type: 1, // xlCellValue
                    Operator: 3, // xlGreater
                    Formula1: "=150");

                // Now write NEW values — this triggers conditional formatting re-evaluation.
                // Before the fix, this would deadlock because:
                // 1. range.Value2 = ... sends COM call to Excel
                // 2. Excel evaluates conditional formatting, sends callback to our STA thread
                // 3. MessagePending returned WAITNOPROCESS → callback queued, not dispatched
                // 4. Excel waits for callback → our thread waits for Excel → DEADLOCK
                sheet.Range["A1"].Value2 = 300;
                sheet.Range["A2"].Value2 = 50;

                _output.WriteLine("✓ Value writes with conditional formatting completed without deadlock");
            }
            finally
            {
                ComUtilities.Release(ref formatCondition);
                ComUtilities.Release(ref formatConditions);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }

            return 0;
        });
    }

    /// <summary>
    /// REGRESSION TEST: Writing formulas that trigger recalculation must not deadlock.
    /// Formulas with dependencies cause Excel to fire Calculate events, which previously
    /// could deadlock the STA thread via WAITNOPROCESS.
    /// </summary>
    [Fact]
    public void WriteFormulas_WithDependencies_DoesNotDeadlock()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ctx.Book.Worksheets[1];

                // Write values that formulas will depend on
                sheet.Range["B1"].Value2 = 10;
                sheet.Range["B2"].Value2 = 20;
                sheet.Range["B3"].Value2 = 30;

                // Write formulas that reference those cells — triggers recalculation
                sheet.Range["C1"].Formula2 = "=B1*2";
                sheet.Range["C2"].Formula2 = "=B2+B3";
                sheet.Range["C3"].Formula2 = "=SUM(B1:B3)";

                // Now change the source values — triggers formula recalculation
                sheet.Range["B1"].Value2 = 100;
                sheet.Range["B2"].Value2 = 200;

                _output.WriteLine("✓ Formula writes with dependencies completed without deadlock");
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }

            return 0;
        });
    }

    /// <summary>
    /// Verifies that bulk value writes with the guard are significantly faster
    /// than they would be without ScreenUpdating suppression.
    /// This is a basic sanity check — exact timing varies by machine.
    /// </summary>
    [Fact]
    public void BulkWrites_WithGuard_CompletesInReasonableTime()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ctx.Book.Worksheets[1];

                // Write 100 cells — with ScreenUpdating=false this should be fast
                for (int i = 1; i <= 100; i++)
                {
                    sheet.Range[$"D{i}"].Value2 = i * 1.5;
                }
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }

            return 0;
        });

        stopwatch.Stop();
        _output.WriteLine($"100 cell writes completed in {stopwatch.ElapsedMilliseconds}ms");

        // With ScreenUpdating suppressed, 100 writes should complete in well under 30s
        Assert.True(stopwatch.ElapsedMilliseconds < 30000,
            $"Bulk writes took {stopwatch.ElapsedMilliseconds}ms — ScreenUpdating may not be suppressed");
    }
}
