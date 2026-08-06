using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;

namespace Sbroenne.ExcelMcp.Benchmarks.Scenarios;

internal sealed class VectorizedWritesScenario : IBenchmarkScenario
{
    private static readonly int[] RangeCellCounts = [100, 1_000, 10_000];
    private static readonly int[] TableRowCounts = [10, 100, 500];
    private const int ColumnCount = 10;

    public string PlanId => "05";

    public string Name => "vectorized-writes";

    public Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken)
    {
        var rangeMaster = context.CreateWorkingPath("vector-range-master");
        var tableMaster = context.CreateWorkingPath("vector-table-master");
        BenchmarkContext.CreateDataWorkbook(rangeMaster, rows: 1_100, columns: ColumnCount);
        BenchmarkContext.CreateDataWorkbook(tableMaster, rows: 2, columns: ColumnCount, includeTable: true);
        var rangeCommands = new RangeCommands();
        var tableCommands = new TableCommands();
        var observations = new List<BenchmarkObservation>();
        var rangeValuesEqual = true;
        var tableValuesEqual = true;
        var formulasEqual = true;
        var multiAreaRejected = true;

        foreach (var cellCount in RangeCellCounts)
        {
            var rows = cellCount / ColumnCount;
            var values = CreateNumericValues(rows, ColumnCount, seed: cellCount);
            var payloadBytes = JsonSerializer.SerializeToUtf8Bytes(values).Length;
            var total = context.Configuration.Warmups + context.Configuration.Iterations;
            for (var iteration = 0; iteration < total; iteration++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var workbookPath = context.CopyWorkbook(rangeMaster, $"range-{cellCount}");
                using var batch = ExcelSession.BeginBatch(context.Configuration.ShowExcel, operationTimeout: null, workbookPath);
                var address = $"A2:J{rows + 1}";
                var memoryBefore = BenchmarkContext.GetWorkingSetBytes();
                var started = Stopwatch.GetTimestamp();
                var result = rangeCommands.SetValues(batch, "Data", address, values);
                var elapsed = BenchmarkContext.ElapsedMilliseconds(started);
                var roundTrip = rangeCommands.GetValues(batch, "Data", address);
                var equal = result.Success && NumericValuesEqual(values, roundTrip.Values);
                rangeValuesEqual &= equal;

                if (iteration >= context.Configuration.Warmups)
                {
                    observations.Add(new BenchmarkObservation(
                        iteration - context.Configuration.Warmups,
                        $"contiguous-range-{cellCount}",
                        equal,
                        result.ErrorMessage,
                        new Dictionary<string, double>
                        {
                            ["write_latency_ms"] = elapsed,
                            ["cells_per_second"] = cellCount / (elapsed / 1000d),
                            ["rows_per_second"] = rows / (elapsed / 1000d),
                            ["payload_bytes"] = payloadBytes,
                            ["working_set_delta_bytes"] = BenchmarkContext.GetWorkingSetBytes() - memoryBefore
                        },
                        new Dictionary<string, string>
                        {
                            ["write_path"] = "Range.Value2-2d-array",
                            ["cells"] = cellCount.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        },
                        "round-trip-verified"));
                }
            }
        }

        foreach (var rowCount in TableRowCounts)
        {
            var rows = CreateNumericValues(rowCount, ColumnCount, seed: rowCount * 100);
            var payloadBytes = JsonSerializer.SerializeToUtf8Bytes(rows).Length;
            var total = context.Configuration.Warmups + context.Configuration.Iterations;
            for (var iteration = 0; iteration < total; iteration++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var workbookPath = context.CopyWorkbook(tableMaster, $"table-{rowCount}");
                using var batch = ExcelSession.BeginBatch(context.Configuration.ShowExcel, operationTimeout: null, workbookPath);
                var memoryBefore = BenchmarkContext.GetWorkingSetBytes();
                var started = Stopwatch.GetTimestamp();
                var result = tableCommands.Append(batch, "BenchmarkTable", rows);
                var elapsed = BenchmarkContext.ElapsedMilliseconds(started);
                var tableData = tableCommands.GetData(batch, "BenchmarkTable", visibleOnly: false);
                var appended = tableData.Data.TakeLast(rowCount).ToArray();
                var equal = result.Success && tableData.RowCount == rowCount + 1 && NumericValuesEqual(rows, appended);
                tableValuesEqual &= equal;

                if (iteration >= context.Configuration.Warmups)
                {
                    observations.Add(new BenchmarkObservation(
                        iteration - context.Configuration.Warmups,
                        $"table-append-{rowCount}",
                        equal,
                        result.ErrorMessage,
                        new Dictionary<string, double>
                        {
                            ["write_latency_ms"] = elapsed,
                            ["cells_per_second"] = rowCount * ColumnCount / (elapsed / 1000d),
                            ["rows_per_second"] = rowCount / (elapsed / 1000d),
                            ["payload_bytes"] = payloadBytes,
                            ["working_set_delta_bytes"] = BenchmarkContext.GetWorkingSetBytes() - memoryBefore
                        },
                        new Dictionary<string, string>
                        {
                            ["write_path"] = "TableCommands.Append-vectorized",
                            ["rows"] = rowCount.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        },
                        "round-trip-verified"));
                }
            }
        }

        formulasEqual = VerifyFormulaRoundTrip(context, rangeMaster, rangeCommands);
        multiAreaRejected = VerifyMultiAreaFailsClosed(context, rangeMaster, rangeCommands);

        var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == PlanId);
        return Task.FromResult(ScenarioResult.Create(
            PlanId,
            Name,
            plan.Title,
            observations,
            [
                new BenchmarkInvariant("round_trip_values_equal", rangeValuesEqual, $"Every contiguous range matched its source matrix: {rangeValuesEqual}"),
                new BenchmarkInvariant("round_trip_formulas_equal", formulasEqual, $"A 100-cell formula matrix round-tripped exactly: {formulasEqual}"),
                new BenchmarkInvariant("table_shape_equal", tableValuesEqual, $"Every append had the expected row/column shape and values: {tableValuesEqual}"),
                new BenchmarkInvariant("no_silent_multi_area_reorder", multiAreaRejected, $"A disjoint target was rejected before an ambiguous write: {multiAreaRejected}")
            ],
            $"Contiguous writes of {string.Join(", ", RangeCellCounts)} cells and table appends of {string.Join(", ", TableRowCounts)} rows × {ColumnCount} columns.",
            ["Range writes are already vectorized; the table-append cases isolate the known cell-by-cell COM hotspot."]));
    }

    private static List<List<object?>> CreateNumericValues(int rows, int columns, int seed)
    {
        var result = new List<List<object?>>(rows);
        for (var row = 0; row < rows; row++)
        {
            var values = new List<object?>(columns);
            for (var column = 0; column < columns; column++)
            {
                values.Add(seed + (row * columns) + column + 0.25d);
            }

            result.Add(values);
        }

        return result;
    }

    private static bool NumericValuesEqual(
        List<List<object?>> expected,
        IReadOnlyList<List<object?>> actual)
    {
        if (expected.Count != actual.Count)
        {
            return false;
        }

        for (var row = 0; row < expected.Count; row++)
        {
            if (expected[row].Count != actual[row].Count)
            {
                return false;
            }

            for (var column = 0; column < expected[row].Count; column++)
            {
                var expectedValue = Convert.ToDouble(expected[row][column], System.Globalization.CultureInfo.InvariantCulture);
                var actualValue = Convert.ToDouble(actual[row][column], System.Globalization.CultureInfo.InvariantCulture);
                if (Math.Abs(expectedValue - actualValue) > 0.000001)
                {
                    return false;
                }
            }
        }

        return true;
    }

    private static bool VerifyFormulaRoundTrip(
        BenchmarkContext context,
        string masterPath,
        RangeCommands rangeCommands)
    {
        var workbookPath = context.CopyWorkbook(masterPath, "formula-round-trip");
        using var batch = ExcelSession.BeginBatch(context.Configuration.ShowExcel, operationTimeout: null, workbookPath);
        var formulas = Enumerable.Range(0, 10)
            .Select(_ => Enumerable.Range(0, 10).Select(_ => "=ROW()+COLUMN()").ToList())
            .ToList();
        var result = rangeCommands.SetFormulas(batch, "Data", "A2:J11", formulas);
        var roundTrip = rangeCommands.GetFormulas(batch, "Data", "A2:J11");
        return result.Success && formulas.SelectMany(row => row).SequenceEqual(roundTrip.Formulas.SelectMany(row => row));
    }

    private static bool VerifyMultiAreaFailsClosed(
        BenchmarkContext context,
        string masterPath,
        RangeCommands rangeCommands)
    {
        var workbookPath = context.CopyWorkbook(masterPath, "multi-area");
        using var batch = ExcelSession.BeginBatch(context.Configuration.ShowExcel, operationTimeout: null, workbookPath);

        var valuesRejected = IsRejected(() => rangeCommands.SetValues(
            batch,
            "Data",
            "A2:A3,C2:C3",
            [
                [11d, 12d],
                [13d, 14d]
            ]));
        var formulasRejected = IsRejected(() => rangeCommands.SetFormulas(
            batch,
            "Data",
            "A2:A3,C2:C3",
            [
                ["=1", "=2"],
                ["=3", "=4"]
            ]));

        return valuesRejected && formulasRejected;

        static bool IsRejected(Action write)
        {
            try
            {
                write();
                // A provider that silently accepts the unsupported SAFEARRAY shape
                // must still fail the gate rather than treating success as rejection.
                return false;
            }
            catch (ArgumentException)
            {
                return true;
            }
            catch (InvalidOperationException)
            {
                return true;
            }
        }
    }
}
