using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.Benchmarks.Scenarios;

internal sealed class ReadFastPathScenario : IBenchmarkScenario
{
    private const int ReadRows = 1_000;
    private const int ReadColumns = 10;

    public string PlanId => "06";

    public string Name => "read-fast-path";

    public async Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken)
    {
        var csvPath = context.CreateWorkingPath("refresh-source", ".csv");
        File.WriteAllText(csvPath, "Value\n0\n");
        var masterPath = context.CreateWorkingPath("read-master");
        BenchmarkContext.CreateDataWorkbook(masterPath, rows: ReadRows + 1, columns: ReadColumns);
        AddFormulaAndRefreshTable(masterPath, csvPath);

        var observations = new List<BenchmarkObservation>();
        var roundTripsEqual = true;
        var ownWriteFresh = true;
        var directEditFresh = true;
        var recalculationFresh = true;
        var refreshFresh = true;
        var total = context.Configuration.Warmups + context.Configuration.Iterations;

        for (var iteration = 0; iteration < total; iteration++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var workbookPath = context.CopyWorkbook(masterPath, "read");
            var safetyRoot = context.CreateSafetyRoot("read");
            using var service = new ExcelMcpService(safetyRoot);
            var sessionId = await ServiceBenchmarkHelpers.CreateSessionAsync(
                service,
                workbookPath,
                context.Configuration.ShowExcel);
            await ServiceBenchmarkHelpers.ConfigureSafetyAsync(
                service,
                sessionId,
                reviewMode: "off",
                checkpointMode: "off",
                journalMode: "off",
                verificationMode: "off");

            var readRequest = new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new
                {
                    sheetName = "Data",
                    rangeAddress = $"A2:J{ReadRows + 1}"
                }, ServiceProtocol.JsonOptions),
                Source = "benchmark"
            };

            var coldStarted = Stopwatch.GetTimestamp();
            var cold = await service.ProcessAsync(readRequest);
            var coldLatency = BenchmarkContext.ElapsedMilliseconds(coldStarted);
            ServiceBenchmarkHelpers.EnsureSuccess(cold, "cold range read");
            var coldCorrect = VerifyReadPayload(cold.Result, ReadRows, ReadColumns);

            var warmStarted = Stopwatch.GetTimestamp();
            var warm = await service.ProcessAsync(readRequest);
            var warmLatency = BenchmarkContext.ElapsedMilliseconds(warmStarted);
            ServiceBenchmarkHelpers.EnsureSuccess(warm, "warm range read");
            var warmCorrect = VerifyReadPayload(warm.Result, ReadRows, ReadColumns);
            roundTripsEqual &= coldCorrect && warmCorrect;

            var ownWriteValue = 200_000d + iteration;
            var write = await service.ProcessAsync(new ServiceRequest
            {
                Command = "range.set-values",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new
                {
                    sheetName = "Data",
                    rangeAddress = "A2",
                    values = new object?[][] { [ownWriteValue] }
                }, ServiceProtocol.JsonOptions),
                Source = "benchmark"
            });
            ServiceBenchmarkHelpers.EnsureSuccess(write, "own write before read");
            ownWriteFresh &= await ReadSingleNumberAsync(service, sessionId, "Data", "A2") == ownWriteValue;

            var batch = service.SessionManager.GetSession(sessionId)
                ?? throw new InvalidOperationException("Read benchmark session disappeared.");
            var directEditValue = 300_000d + iteration;
            batch.Execute((excelContext, _) =>
            {
                WriteDirectCell(excelContext, "Data", "B2", directEditValue);
                return 0;
            }, cancellationToken);
            directEditFresh &= await ReadSingleNumberAsync(service, sessionId, "Data", "B2") == directEditValue;

            var formulaInput = 400_000d + iteration;
            batch.Execute((excelContext, _) =>
            {
                WriteDirectCell(excelContext, "Data", "A2", formulaInput);
                excelContext.App.Calculate();
                return 0;
            }, cancellationToken);
            recalculationFresh &= await ReadSingleNumberAsync(service, sessionId, "Data", "K2") == formulaInput * 2;

            var refreshValue = 500_000 + iteration;
            File.WriteAllText(csvPath, $"Value{Environment.NewLine}{refreshValue.ToString(CultureInfo.InvariantCulture)}{Environment.NewLine}");
            var refreshStarted = Stopwatch.GetTimestamp();
            Refresh(batch, cancellationToken);
            var refreshedValue = await ReadSingleNumberAsync(service, sessionId, "RefreshData", "A2");
            var refreshLatency = BenchmarkContext.ElapsedMilliseconds(refreshStarted);
            var refreshCorrect = Math.Abs(refreshedValue - refreshValue) < 0.000001;
            refreshFresh &= refreshCorrect;

            if (iteration >= context.Configuration.Warmups)
            {
                var payloadBytes = ServiceBenchmarkHelpers.SerializedPayloadBytes(readRequest, warm);
                observations.Add(new BenchmarkObservation(
                    iteration - context.Configuration.Warmups,
                    "service-range-read-and-local-csv-refresh",
                    coldCorrect && warmCorrect && refreshCorrect,
                    null,
                    new Dictionary<string, double>
                    {
                        ["cold_read_ms"] = coldLatency,
                        ["warm_read_ms"] = warmLatency,
                        ["refresh_to_consistent_read_ms"] = refreshLatency,
                        ["payload_bytes"] = payloadBytes,
                        ["token_estimate"] = BenchmarkContext.EstimateTokensFromUtf8Bytes(payloadBytes)
                    },
                    new Dictionary<string, string>
                    {
                        ["read_cells"] = (ReadRows * ReadColumns).ToString(CultureInfo.InvariantCulture),
                        ["refresh_source"] = "local-csv-querytable-backgroundQuery=false",
                        ["token_measurement"] = "ceil(utf8-request-response-bytes/4)"
                    },
                    "consistent"));
            }

            await ServiceBenchmarkHelpers.CloseSessionAsync(service, sessionId);
        }

        var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == PlanId);
        return ScenarioResult.Create(
            PlanId,
            Name,
            plan.Title,
            observations,
            [
                new BenchmarkInvariant("round_trip_values_equal", roundTripsEqual, $"Cold and warm 10,000-cell reads returned expected shape/literals: {roundTripsEqual}"),
                new BenchmarkInvariant("no_stale_read_after_write", ownWriteFresh && directEditFresh && recalculationFresh, $"Own write fresh: {ownWriteFresh}; direct edit fresh: {directEditFresh}; recalculation fresh: {recalculationFresh}"),
                new BenchmarkInvariant("no_stale_read_after_refresh", refreshFresh, $"Every local CSV refresh returned the newly written source value: {refreshFresh}"),
                new BenchmarkInvariant("refresh_result_consistent", refreshFresh, $"Refresh completed synchronously and first read was correct: {refreshFresh}")
            ],
            "Cold and warm 10,000-cell service reads, direct-edit and recalculation invalidation checks, and an offline local-CSV QueryTable refresh.",
            ["The token figure measures this fixed service request/response payload; it is an estimate, while bytes are exact."]);
    }

    private static bool VerifyReadPayload(string? result, int expectedRows, int expectedColumns)
    {
        if (string.IsNullOrWhiteSpace(result))
        {
            return false;
        }

        using var document = JsonDocument.Parse(result);
        var root = document.RootElement;
        if (root.GetProperty("rowCount").GetInt32() != expectedRows ||
            root.GetProperty("columnCount").GetInt32() != expectedColumns)
        {
            return false;
        }

        var values = root.GetProperty("values");
        if (values.GetArrayLength() != expectedRows)
        {
            return false;
        }

        var rowIndex = 0;
        foreach (var row in values.EnumerateArray())
        {
            if (row.GetArrayLength() != expectedColumns)
            {
                return false;
            }

            var columnIndex = 0;
            foreach (var value in row.EnumerateArray())
            {
                var expected = ((rowIndex + 1) * 1000d) + columnIndex + 1;
                if (value.ValueKind != JsonValueKind.Number || Math.Abs(value.GetDouble() - expected) > 0.000001)
                {
                    return false;
                }

                columnIndex++;
            }

            rowIndex++;
        }

        return true;
    }

    private static async Task<double> ReadSingleNumberAsync(
        ExcelMcpService service,
        string sessionId,
        string sheetName,
        string rangeAddress)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { sheetName, rangeAddress }, ServiceProtocol.JsonOptions),
            Source = "benchmark"
        });
        ServiceBenchmarkHelpers.EnsureSuccess(response, $"read {sheetName}!{rangeAddress}");
        using var document = JsonDocument.Parse(response.Result!);
        return document.RootElement.GetProperty("values")[0][0].GetDouble();
    }

    private static void AddFormulaAndRefreshTable(string workbookPath, string csvPath)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((excelContext, _) =>
        {
            dynamic? dataSheet = null;
            dynamic? formulaRange = null;
            dynamic? sheets = null;
            dynamic? refreshSheet = null;
            dynamic? destination = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;
            try
            {
                dataSheet = excelContext.Book.Worksheets["Data"];
                formulaRange = dataSheet.Range[$"K2:K{ReadRows + 1}"];
                formulaRange.FormulaR1C1 = "=RC[-10]*2";

                sheets = excelContext.Book.Worksheets;
                refreshSheet = sheets.Add();
                refreshSheet.Name = "RefreshData";
                destination = refreshSheet.Range["A1"];
                queryTables = refreshSheet.QueryTables;
                queryTable = queryTables.Add($"TEXT;{csvPath}", destination);
                queryTable.Name = "BenchmarkRefresh";
                queryTable.TextFileParseType = 1;
                queryTable.TextFileCommaDelimiter = true;
                queryTable.TextFileStartRow = 1;
                queryTable.BackgroundQuery = false;
                queryTable.Refresh();
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref refreshSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref formulaRange);
                ComUtilities.Release(ref dataSheet);
            }
        });
        batch.Save();
    }

    private static void WriteDirectCell(ExcelContext context, string sheetName, string address, double value)
    {
        dynamic? sheet = null;
        dynamic? cell = null;
        try
        {
            sheet = context.Book.Worksheets[sheetName];
            cell = sheet.Range[address];
            cell.Value2 = value;
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref sheet);
        }
    }

    private static void Refresh(IExcelBatch batch, CancellationToken cancellationToken)
    {
        batch.Execute((excelContext, _) =>
        {
            dynamic? sheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;
            try
            {
                sheet = excelContext.Book.Worksheets["RefreshData"];
                queryTables = sheet.QueryTables;
                queryTable = queryTables.Item(1);
                queryTable.BackgroundQuery = false;
                queryTable.Refresh();
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref sheet);
            }
        }, cancellationToken);
    }
}
