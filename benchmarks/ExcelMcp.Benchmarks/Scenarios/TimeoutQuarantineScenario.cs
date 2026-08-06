using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Benchmarks.Scenarios;

internal sealed class TimeoutQuarantineScenario : IBenchmarkScenario
{
    private static readonly TimeSpan OperationTimeout = TimeSpan.FromMilliseconds(350);
    private const int BlockMilliseconds = 1_500;

    public string PlanId => "01";

    public string Name => "timeout-quarantine";

    public Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken)
    {
        var observations = new List<BenchmarkObservation>();
        var timedOutEveryTime = true;
        var failFastEveryTime = true;
        var exitedEveryTime = true;
        var lateWriteCount = 0;

        var total = context.Configuration.Warmups + context.Configuration.ReliabilityIterations;
        for (var index = 0; index < total; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var measuredIteration = index - context.Configuration.Warmups;
            var workbookPath = context.CreateWorkingPath("timeout");
            BenchmarkContext.CreateDataWorkbook(workbookPath, rows: 2, columns: 1);
            var workingSetBefore = BenchmarkContext.GetWorkingSetBytes();
            var returnLatency = 0d;
            var cleanupLatency = 0d;
            var failFastLatency = 0d;
            var timedOut = false;
            var failFast = false;
            var processExited = false;
            var error = default(string);
            int? processId = null;
            IExcelBatch? batch = null;

            try
            {
                batch = ExcelSession.BeginBatch(
                    context.Configuration.ShowExcel,
                    operationTimeout: null,
                    workbookPath);
                processId = batch.ExcelProcessId;

                var operationStarted = Stopwatch.GetTimestamp();
                using var operationCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                operationCancellation.CancelAfter(OperationTimeout);
                try
                {
                    batch.Execute((excelContext, _) =>
                    {
                        ComUtilities.KernelSleep(BlockMilliseconds);
                        dynamic? sheet = null;
                        dynamic? cell = null;
                        try
                        {
                            sheet = excelContext.Book.Worksheets[1];
                            cell = sheet.Range["A2"];
                            cell.Value2 = 999d;
                            excelContext.Book.Save();
                        }
                        finally
                        {
                            ComUtilities.Release(ref cell);
                            ComUtilities.Release(ref sheet);
                        }

                        return 0;
                    }, operationCancellation.Token);
                }
                catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
                {
                    timedOut = true;
                }
                finally
                {
                    returnLatency = BenchmarkContext.ElapsedMilliseconds(operationStarted);
                }

                var failFastStarted = Stopwatch.GetTimestamp();
                try
                {
                    batch.Execute(static (_, _) => 0, cancellationToken);
                }
                catch (InvalidOperationException)
                {
                    failFast = true;
                }
                catch (TimeoutException)
                {
                    failFast = true;
                }
                finally
                {
                    failFastLatency = BenchmarkContext.ElapsedMilliseconds(failFastStarted);
                }
            }
            catch (Exception exception)
            {
                error = $"{exception.GetType().Name}: {exception.Message}";
            }
            finally
            {
                var cleanupStarted = Stopwatch.GetTimestamp();
                batch?.Dispose();
                if (processId.HasValue)
                {
                    processExited = BenchmarkContext.WaitForProcessExit(
                        processId.Value,
                        TimeSpan.FromSeconds(15),
                        out _);
                    cleanupLatency = BenchmarkContext.ElapsedMilliseconds(cleanupStarted);
                }
            }

            var persistedValue = ReadCell(workbookPath, context.Configuration.ShowExcel);
            var noLateWrite = persistedValue is not 999d;
            timedOutEveryTime &= timedOut;
            failFastEveryTime &= failFast;
            exitedEveryTime &= processExited;
            if (!noLateWrite)
            {
                lateWriteCount++;
            }

            if (index >= context.Configuration.Warmups)
            {
                var success = error is null && timedOut && failFast && processExited && noLateWrite;
                observations.Add(new BenchmarkObservation(
                    measuredIteration,
                    "post-dispatch-timeout",
                    success,
                    error,
                    new Dictionary<string, double>
                    {
                        ["return_latency_ms"] = returnLatency,
                        ["cleanup_latency_ms"] = cleanupLatency,
                        ["fail_fast_latency_ms"] = failFastLatency,
                        ["orphan_process_count"] = processExited ? 0 : 1,
                        ["working_set_delta_bytes"] = BenchmarkContext.GetWorkingSetBytes() - workingSetBefore
                    },
                    new Dictionary<string, string>
                    {
                        ["caller_deadline_ms"] = OperationTimeout.TotalMilliseconds.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        ["blocked_operation_ms"] = BlockMilliseconds.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        ["excel_process_id"] = processId?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "not-captured"
                    },
                    timedOut ? "abortedUnknown" : "unexpected"));
            }
        }

        var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == PlanId);
        var invariants = new List<BenchmarkInvariant>
        {
            new("no_post_timeout_write", lateWriteCount == 0, $"Late persisted writes: {lateWriteCount}"),
            new("outcome_unknown_not_success", timedOutEveryTime, $"Every run timed out: {timedOutEveryTime}"),
            new("session_unusable_after_timeout", failFastEveryTime, $"Every poisoned session rejected later work: {failFastEveryTime}"),
            new("no_owned_excel_orphan", exitedEveryTime, $"Every captured owned Excel PID exited: {exitedEveryTime}")
        };

        return Task.FromResult(ScenarioResult.Create(
            PlanId,
            Name,
            plan.Title,
            observations,
            invariants,
            $"{context.Configuration.ReliabilityIterations} isolated real-Excel sessions; after normal startup, each operation exceeded a {OperationTimeout.TotalMilliseconds:0} ms caller deadline and then attempted a delayed write.",
            ["A COM call cannot be safely pre-empted; this benchmark measures containment and explicit outcome uncertainty."]));
    }

    private static object? ReadCell(string workbookPath, bool showExcel)
    {
        using var batch = ExcelSession.BeginBatch(showExcel, operationTimeout: null, workbookPath);
        return batch.Execute((context, _) =>
        {
            dynamic? sheet = null;
            dynamic? cell = null;
            try
            {
                sheet = context.Book.Worksheets[1];
                cell = sheet.Range["A2"];
                return cell.Value2;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
