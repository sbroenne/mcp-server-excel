using System.Diagnostics;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

/// <summary>
/// Integration tests for ExcelWriteGuard — the structural COM safety mechanism
/// integrated into ExcelBatch.Execute().
///
/// Verifies ScreenUpdating suppression and restoration without changing events
/// or calculation. Direct guard controls run on their own STA so verification
/// occurs after disposal, outside any Execute guard.
///
/// Regression controls for message-pump deadlocks:
/// - Range writes triggering Calculate callbacks → WAITNOPROCESS deadlock
/// - Conditional formatting operations with dependent formulas
/// - Bulk writes without ScreenUpdating suppression
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelWriteGuard")]
[Collection("Sequential")]
[Trait("RequiresExcel", "true")]
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
            File.Delete(_testFileCopy);
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
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public void Guard_RestoresState_AfterDisposal(bool originalScreenUpdating, bool throwInside)
    {
        RunStandaloneGuardControl(app =>
        {
            app.ScreenUpdating = originalScreenUpdating;
            app.EnableEvents = false;
            app.Calculation = Excel.XlCalculation.xlCalculationManual;
            var error = Record.Exception(() =>
            {
                using var guard = new ExcelWriteGuard(app);
                Assert.False(app.ScreenUpdating);
                Assert.False(app.EnableEvents);
                Assert.Equal(Excel.XlCalculation.xlCalculationManual, app.Calculation);
                if (throwInside) throw new InvalidOperationException("guard-control");
            });
            if (throwInside)
                Assert.Equal("guard-control", Assert.IsType<InvalidOperationException>(error).Message);
            else
                Assert.Null(error);
            Assert.Equal(originalScreenUpdating, app.ScreenUpdating);
            Assert.False(app.EnableEvents);
            Assert.Equal(Excel.XlCalculation.xlCalculationManual, app.Calculation);
        });
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
            using (var innerGuard = new ExcelWriteGuard(ctx.App))
            {
                innerScreenUpdating = ctx.App.ScreenUpdating;
                Assert.False(innerScreenUpdating);
            }
            Assert.False(ctx.App.ScreenUpdating);
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
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.FormatConditions? formatConditions = null;
            Excel.FormatCondition? formatCondition = null;

            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];

                // Set up: write initial values
                SetCellValue(sheet, "A1", 100);
                SetCellValue(sheet, "A2", 200);

                // Add conditional formatting rule on A1:A2
                range = sheet.Range["A1:A2"];
                formatConditions = range.FormatConditions;
                formatCondition = (Excel.FormatCondition)formatConditions.Add(
                    Excel.XlFormatConditionType.xlCellValue,
                    Excel.XlFormatConditionOperator.xlGreater, "=150");

                // Now write NEW values — this triggers conditional formatting re-evaluation.
                // Before the fix, this would deadlock because:
                // 1. range.Value2 = ... sends COM call to Excel
                // 2. Excel evaluates conditional formatting, sends callback to our STA thread
                // 3. MessagePending returned WAITNOPROCESS → callback queued, not dispatched
                // 4. Excel waits for callback → our thread waits for Excel → DEADLOCK
                SetCellValue(sheet, "A1", 300);
                SetCellValue(sheet, "A2", 50);
                var values = Assert.IsType<object[,]>(range.Value2);
                Assert.Equal(300d, values[1, 1]);
                Assert.Equal(50d, values[2, 1]);
                Assert.Equal(1, formatConditions.Count);
                Assert.Equal((int)Excel.XlFormatConditionOperator.xlGreater, formatCondition.Operator);
                Assert.Equal("=150", formatCondition.Formula1);

                _output.WriteLine("✓ Value writes with conditional formatting completed without deadlock");
            }
            finally
            {
                ComUtilities.Release(ref formatCondition);
                ComUtilities.Release(ref formatConditions);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
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
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? formulas = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];

                // Write values that formulas will depend on
                SetCellValue(sheet, "B1", 10);
                SetCellValue(sheet, "B2", 20);
                SetCellValue(sheet, "B3", 30);

                // Write formulas that reference those cells — triggers recalculation
                SetCellFormula(sheet, "C1", "=B1*2");
                SetCellFormula(sheet, "C2", "=B2+B3");
                SetCellFormula(sheet, "C3", "=SUM(B1:B3)");

                // Now change the source values — triggers formula recalculation
                SetCellValue(sheet, "B1", 100);
                SetCellValue(sheet, "B2", 200);
                ctx.App.Calculate();
                formulas = sheet.Range["C1:C3"];
                var values = Assert.IsType<object[,]>(formulas.Value2);
                Assert.Equal(200d, values[1, 1]);
                Assert.Equal(230d, values[2, 1]);
                Assert.Equal(330d, values[3, 1]);

                _output.WriteLine("✓ Formula writes with dependencies completed without deadlock");
            }
            finally
            {
                ComUtilities.Release(ref formulas);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }

            return 0;
        });
    }

    /// <summary>
    /// Verifies complete bulk-write results within a bounded operation time.
    /// </summary>
    [Fact]
    public void BulkWrites_WithGuard_CompletesInReasonableTime()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? written = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];

                // Write 100 cells — with ScreenUpdating=false this should be fast
                for (int i = 1; i <= 100; i++)
                {
                    SetCellValue(sheet, $"D{i}", i * 1.5);
                }
                written = sheet.Range["D1:D100"];
                var values = Assert.IsType<object[,]>(written.Value2);
                for (int row = 1; row <= 100; row++)
                    Assert.Equal(row * 1.5, values[row, 1]);
            }
            finally
            {
                ComUtilities.Release(ref written);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }

            return 0;
        });

        stopwatch.Stop();
        _output.WriteLine($"100 cell writes completed in {stopwatch.ElapsedMilliseconds}ms");

        // With ScreenUpdating suppressed, 100 writes should complete in well under 30s
        Assert.True(stopwatch.ElapsedMilliseconds < 30000,
            $"Bulk writes took {stopwatch.ElapsedMilliseconds}ms — ScreenUpdating may not be suppressed");
    }

    private static void SetCellValue(Excel.Worksheet sheet, string address, object value)
    {
        Excel.Range? cell = null;
        try
        {
            cell = sheet.Range[address];
            cell.Value2 = value;
        }
        finally
        {
            ComUtilities.Release(ref cell);
        }
    }

    private static void SetCellFormula(Excel.Worksheet sheet, string address, string formula)
    {
        Excel.Range? cell = null;
        try
        {
            cell = sheet.Range[address];
            cell.Formula2 = formula;
        }
        finally
        {
            ComUtilities.Release(ref cell);
        }
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

    private static void RunStandaloneGuardControl(Action<Excel.Application> assertion)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            Excel.Application? app = null;
            Excel.Workbooks? books = null;
            Excel.Workbook? book = null;
            ExcelProcessIdentity? identity = null;
            var failures = new List<Exception>();
            try
            {
                OleMessageFilter.Register();
                app = new Excel.Application { DisplayAlerts = false };
                Assert.NotEqual(0u, GetWindowThreadProcessId(new IntPtr(app.Hwnd), out var pid));
                Assert.NotEqual(0u, pid);
                using var process = Process.GetProcessById(checked((int)pid));
                identity = new ExcelProcessIdentity(process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());
                SessionManager.TrackExcelProcess(identity.Value);
                books = app.Workbooks;
                book = books.Add();
                assertion(app);
            }
            catch (Exception ex)
            {
                failures.Add(ex);
            }
            finally
            {
                CaptureCleanup(() => ComUtilities.Release(ref books));
                CaptureCleanup(() => ExcelShutdownService.CloseAndQuit(book, app, save: false));
                CaptureCleanup(OleMessageFilter.Revoke);
                if (identity is { } owned)
                {
                    CaptureCleanup(() =>
                    {
                        var exited = SpinWait.SpinUntil(() => OwnedProcessGuard.TryConfirmExited(owned),
                            TimeSpan.FromSeconds(15));
                        if (!exited)
                            OwnedProcessGuard.TryTerminate(owned, TimeSpan.Zero, TimeSpan.FromSeconds(5), out _);
                        Assert.True(exited, "The guard control's owned Excel process survived cleanup.");
                        SessionManager.UntrackExcelProcess(owned);
                    });
                }
            }
            if (failures.Count == 0) completion.SetResult();
            else completion.SetException(new AggregateException("Guard control or cleanup failed.", failures));

            void CaptureCleanup(Action cleanup)
            {
                try { cleanup(); }
                catch (Exception ex) { failures.Add(ex); }
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        completion.Task.WaitAsync(TimeSpan.FromSeconds(90)).GetAwaiter().GetResult();
        Assert.True(thread.Join(TimeSpan.FromSeconds(5)));
    }
}
