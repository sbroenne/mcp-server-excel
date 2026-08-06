using System.Collections.Concurrent;
using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Benchmarks.Scenarios;

internal sealed class BoundedWorkbookQueueScenario : IBenchmarkScenario
{
    private static readonly int[] BurstSizes = [4, 16, 32];

    public string PlanId => "02";

    public string Name => "bounded-workbook-queue";

    public async Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken)
    {
        var masterPath = context.CreateWorkingPath("queue-master");
        var independentMasterPath = context.CreateWorkingPath("queue-independent-master");
        BenchmarkContext.CreateDataWorkbook(masterPath, rows: 64, columns: 1);
        BenchmarkContext.CreateDataWorkbook(independentMasterPath, rows: 2, columns: 1);

        var observations = new List<BenchmarkObservation>();
        var allOrdered = true;
        var allWritesPresent = true;
        var noDuplicates = true;
        var independentProgress = true;

        foreach (var burstSize in BurstSizes)
        {
            var total = context.Configuration.Warmups + context.Configuration.Iterations;
            for (var index = 0; index < total; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var workbookPath = context.CopyWorkbook(masterPath, $"queue-{burstSize}");
                var independentPath = context.CopyWorkbook(independentMasterPath, "queue-independent");
                using var batch = ExcelSession.BeginBatch(context.Configuration.ShowExcel, operationTimeout: null, workbookPath);
                using var independentBatch = ExcelSession.BeginBatch(context.Configuration.ShowExcel, operationTimeout: null, independentPath);
                using var holdFirstOperation = new ManualResetEventSlim(false);
                using var firstOperationEntered = new ManualResetEventSlim(false);
                var completionOrder = new ConcurrentQueue<int>();
                var queueWaitMilliseconds = new double[burstSize];
                var enqueueSignals = Enumerable.Range(0, burstSize)
                    .Select(_ => new ManualResetEventSlim(false))
                    .ToArray();
                var workingSetBefore = BenchmarkContext.GetWorkingSetBytes();
                var burstStarted = Stopwatch.GetTimestamp();
                var tasks = new List<Task>(burstSize);

                try
                {
                    for (var operationIndex = 0; operationIndex < burstSize; operationIndex++)
                    {
                        var capturedIndex = operationIndex;
                        tasks.Add(Task.Run(() =>
                        {
                            var queuedAt = Stopwatch.GetTimestamp();
                            enqueueSignals[capturedIndex].Set();
                            batch.Execute((excelContext, _) =>
                            {
                                queueWaitMilliseconds[capturedIndex] = BenchmarkContext.ElapsedMilliseconds(queuedAt);
                                if (capturedIndex == 0)
                                {
                                    firstOperationEntered.Set();
                                    holdFirstOperation.Wait(cancellationToken);
                                }

                                WriteIndexedCell(excelContext, capturedIndex + 1, capturedIndex + 10_000);
                                completionOrder.Enqueue(capturedIndex);
                                return 0;
                            }, cancellationToken);
                        }, cancellationToken));

                        if (!enqueueSignals[capturedIndex].Wait(TimeSpan.FromSeconds(5), cancellationToken))
                        {
                            throw new TimeoutException($"Queue caller {capturedIndex} did not start.");
                        }

                        // Give the caller a deterministic chance to enqueue before starting the next one.
                        await Task.Delay(3, cancellationToken);
                    }

                    if (!firstOperationEntered.Wait(TimeSpan.FromSeconds(5), cancellationToken))
                    {
                        throw new TimeoutException("First queue operation did not enter the STA worker.");
                    }

                    await Task.Delay(50, cancellationToken);
                    var workingSetAfterQueue = BenchmarkContext.GetWorkingSetBytes();

                    var independentStarted = Stopwatch.GetTimestamp();
                    var independentTask = Task.Run(() => independentBatch.Execute(static (_, _) => 42, cancellationToken), cancellationToken);
                    var independentCompletedWhileBlocked = await Task.WhenAny(
                        independentTask,
                        Task.Delay(TimeSpan.FromSeconds(2), cancellationToken)) == independentTask;
                    var independentLatency = BenchmarkContext.ElapsedMilliseconds(independentStarted);
                    holdFirstOperation.Set();
                    await Task.WhenAll(tasks);
                    var independentResult = await independentTask.WaitAsync(TimeSpan.FromSeconds(15), cancellationToken);
                    independentProgress &= independentCompletedWhileBlocked && independentResult == 42;
                    var burstCompletion = BenchmarkContext.ElapsedMilliseconds(burstStarted);
                    var order = completionOrder.ToArray();
                    var ordered = order.SequenceEqual(Enumerable.Range(0, burstSize));
                    var writtenValues = ReadIndexedCells(batch, burstSize);
                    var expectedValues = Enumerable.Range(0, burstSize).Select(item => (double)(item + 10_000)).ToArray();
                    var writesPresent = writtenValues.SequenceEqual(expectedValues);
                    var duplicatesAbsent = writtenValues.Distinct().Count() == burstSize;
                    allOrdered &= ordered;
                    allWritesPresent &= writesPresent;
                    noDuplicates &= duplicatesAbsent;

                    if (index >= context.Configuration.Warmups)
                    {
                        var measuredQueueWaits = queueWaitMilliseconds.Order().ToArray();
                        observations.Add(new BenchmarkObservation(
                            index - context.Configuration.Warmups,
                            $"burst-{burstSize}",
                            ordered && writesPresent && duplicatesAbsent && independentCompletedWhileBlocked,
                            null,
                            new Dictionary<string, double>
                            {
                                ["queue_wait_ms"] = Statistics.Percentile(measuredQueueWaits, 0.50),
                                ["queue_wait_p95_ms"] = Statistics.Percentile(measuredQueueWaits, 0.95),
                                ["burst_completion_ms"] = burstCompletion,
                                ["operations_per_second"] = burstSize / (burstCompletion / 1000d),
                                ["working_set_delta_bytes"] = workingSetAfterQueue - workingSetBefore,
                                ["rejected_operation_count"] = 0,
                                ["independent_workbook_latency_ms"] = independentLatency
                            },
                            new Dictionary<string, string>
                            {
                                ["burst_size"] = burstSize.ToString(System.Globalization.CultureInfo.InvariantCulture),
                                ["queue_implementation"] = "bounded-admission-gate-wait-backpressure",
                                ["queue_capacity"] = ExcelBatch.WorkQueueCapacity.ToString(System.Globalization.CultureInfo.InvariantCulture),
                                ["queue_full_mode"] = ExcelBatch.WorkQueueFullMode.ToString()
                            },
                            "completed"));
                    }
                }
                finally
                {
                    holdFirstOperation.Set();
                    foreach (var signal in enqueueSignals)
                    {
                        signal.Dispose();
                    }
                }
            }
        }

        var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == PlanId);
        return ScenarioResult.Create(
            PlanId,
            Name,
            plan.Title,
            observations,
            [
                new BenchmarkInvariant("fifo_order", allOrdered, $"All measured burst completions followed enqueue order: {allOrdered}"),
                new BenchmarkInvariant("no_dropped_mutation", allWritesPresent, $"Every unique target cell contained its expected value: {allWritesPresent}"),
                new BenchmarkInvariant("no_duplicate_mutation", noDuplicates, $"Every burst produced exactly one copy of each value: {noDuplicates}"),
                new BenchmarkInvariant("independent_workbook_progress", independentProgress, $"A second workbook progressed while the first workbook STA was blocked: {independentProgress}")
            ],
            $"Real-Excel bursts of {string.Join(", ", BurstSizes)} same-workbook mutations, with a second workbook proving cross-session progress.",
            [$"Each workbook queue is bounded to {ExcelBatch.WorkQueueCapacity} waiting items; full queues wait for admission and never use a channel drop mode."]);
    }

    private static void WriteIndexedCell(ExcelContext context, int row, double value)
    {
        dynamic? sheet = null;
        dynamic? cell = null;
        try
        {
            sheet = context.Book.Worksheets[1];
            cell = sheet.Cells[row, 1];
            cell.Value2 = value;
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref sheet);
        }
    }

    private static double[] ReadIndexedCells(IExcelBatch batch, int count)
    {
        return batch.Execute((context, _) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = context.Book.Worksheets[1];
                range = sheet.Range[$"A1:A{count}"];
                object raw = range.Value2;
                if (raw is not object[,] values)
                {
                    return [Convert.ToDouble(raw, System.Globalization.CultureInfo.InvariantCulture)];
                }

                var result = new double[count];
                for (var row = 1; row <= count; row++)
                {
                    result[row - 1] = Convert.ToDouble(values[row, 1], System.Globalization.CultureInfo.InvariantCulture);
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
