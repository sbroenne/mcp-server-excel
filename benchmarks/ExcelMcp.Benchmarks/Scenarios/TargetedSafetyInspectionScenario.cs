using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.Benchmarks.Scenarios;

internal sealed class TargetedSafetyInspectionScenario : IBenchmarkScenario
{
    public string PlanId => "03";

    public string Name => "targeted-safety-inspection";

    public async Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken)
    {
        var smallMaster = context.CreateWorkingPath("inspection-small-master");
        var largeMaster = context.CreateWorkingPath("inspection-large-master");
        BenchmarkContext.CreateDataWorkbook(smallMaster, rows: 100, columns: 10, includeTable: true);
        BenchmarkContext.CreateDataWorkbook(largeMaster, rows: 5_000, columns: 20, includeTable: true);
        AddStructuralComplexity(largeMaster);

        var observations = new List<BenchmarkObservation>();
        var exactScopeEveryTime = true;
        var verificationEveryTime = true;
        var staleRangeRejected = true;
        var staleStructureRejected = true;

        foreach (var variant in new[]
        {
            (Name: "small", Path: smallMaster, DeclaredCells: 1_000),
            (Name: "large-complex", Path: largeMaster, DeclaredCells: 100_990)
        })
        {
            var workbookPath = context.CopyWorkbook(variant.Path, $"inspection-{variant.Name}");
            var safetyRoot = context.CreateSafetyRoot($"inspection-{variant.Name}");
            using var service = new ExcelMcpService(safetyRoot);
            var sessionId = await ServiceBenchmarkHelpers.CreateSessionAsync(
                service,
                workbookPath,
                context.Configuration.ShowExcel);
            await ServiceBenchmarkHelpers.ConfigureSafetyAsync(
                service,
                sessionId,
                reviewMode: "required",
                checkpointMode: "off");

            var total = context.Configuration.Warmups + context.Configuration.Iterations;
            for (var index = 0; index < total; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var mutationArgs = JsonSerializer.Serialize(new
                {
                    sheetName = "Data",
                    rangeAddress = "B2",
                    values = new object?[][] { [100_000 + index] }
                }, ServiceProtocol.JsonOptions);
                var reviewRequest = new ServiceRequest
                {
                    Command = "range.set-values",
                    SessionId = sessionId,
                    Args = mutationArgs,
                    ReviewOnly = true,
                    Source = "benchmark"
                };

                var reviewStarted = Stopwatch.GetTimestamp();
                var review = await service.ProcessAsync(reviewRequest);
                var reviewLatency = BenchmarkContext.ElapsedMilliseconds(reviewStarted);
                ServiceBenchmarkHelpers.EnsureSuccess(review, "range.set-values review");
                var reviewId = BenchmarkContext.GetRequiredString(review.Result, "reviewId");
                var exactScope = ReviewHasExactTargetScope(review.Result, "Data", "B2");
                exactScopeEveryTime &= exactScope;

                var executeRequest = new ServiceRequest
                {
                    Command = "range.set-values",
                    SessionId = sessionId,
                    Args = mutationArgs,
                    ReviewId = reviewId,
                    Source = "benchmark"
                };
                var executeStarted = Stopwatch.GetTimestamp();
                var execute = await service.ProcessAsync(executeRequest);
                var verificationLatency = BenchmarkContext.ElapsedMilliseconds(executeStarted);
                ServiceBenchmarkHelpers.EnsureSuccess(execute, "range.set-values execute");
                var verified = HasVerificationStatus(execute.Result, "verified");
                verificationEveryTime &= verified;

                if (index >= context.Configuration.Warmups)
                {
                    var payloadBytes = ServiceBenchmarkHelpers.SerializedPayloadBytes(reviewRequest, review) +
                        ServiceBenchmarkHelpers.SerializedPayloadBytes(executeRequest, execute);
                    observations.Add(new BenchmarkObservation(
                        index - context.Configuration.Warmups,
                        variant.Name,
                        exactScope && verified,
                        exactScope && verified ? null : "Review scope or post-write verification was not exact.",
                        new Dictionary<string, double>
                        {
                            ["review_latency_ms"] = reviewLatency,
                            ["verification_latency_ms"] = verificationLatency,
                            ["inspected_cell_count"] = exactScope ? 1 : double.NaN,
                            ["inspection_corpus_cell_count"] = variant.DeclaredCells,
                            ["payload_bytes"] = payloadBytes
                        },
                        new Dictionary<string, string>
                        {
                            ["workbook_variant"] = variant.Name,
                            ["corpus_cells"] = variant.DeclaredCells.ToString(System.Globalization.CultureInfo.InvariantCulture),
                            ["scope_evidence"] = "exact affected.sheets/ranges/objects response"
                        },
                        verified ? "verified" : "not-fully-verified"));
                }
            }

            var staleRange = await ExerciseStaleRangeReviewAsync(service, sessionId, cancellationToken);
            staleRangeRejected &= staleRange.Rejected;
            observations.Add(new BenchmarkObservation(
                context.Configuration.Iterations,
                $"{variant.Name}-manual-edit-invalidation",
                staleRange.Rejected,
                staleRange.Error,
                new Dictionary<string, double>
                {
                    ["stale_detection_rate"] = staleRange.Rejected ? 1 : 0,
                    ["stale_detection_ms"] = staleRange.LatencyMilliseconds
                },
                Outcome: staleRange.Category));

            var staleStructure = await ExerciseStaleStructureReviewAsync(service, sessionId, cancellationToken);
            staleStructureRejected &= staleStructure.Rejected;
            observations.Add(new BenchmarkObservation(
                context.Configuration.Iterations + 1,
                $"{variant.Name}-structural-invalidation",
                staleStructure.Rejected,
                staleStructure.Error,
                new Dictionary<string, double>
                {
                    ["stale_detection_rate"] = staleStructure.Rejected ? 1 : 0,
                    ["stale_detection_ms"] = staleStructure.LatencyMilliseconds
                },
                Outcome: staleStructure.Category));

            await ServiceBenchmarkHelpers.CloseSessionAsync(service, sessionId);
        }

        var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == PlanId);
        return ScenarioResult.Create(
            PlanId,
            Name,
            plan.Title,
            observations,
            [
                new BenchmarkInvariant("no_stale_authorization", staleRangeRejected && staleStructureRejected, $"Range stale rejected: {staleRangeRejected}; structural stale rejected: {staleStructureRejected}"),
                new BenchmarkInvariant("exact_affected_scope", exactScopeEveryTime, $"Every review reported only Data!$B$2, with no extra sheets, ranges, or objects: {exactScopeEveryTime}"),
                new BenchmarkInvariant("manual_edit_invalidates", staleRangeRejected, $"Direct COM edit invalidated the pending review: {staleRangeRejected}"),
                new BenchmarkInvariant("structural_change_invalidates", staleStructureRejected, $"Direct worksheet creation invalidated the pending review: {staleStructureRejected}"),
                new BenchmarkInvariant("post_write_verification", verificationEveryTime, $"Every measured mutation returned verified status: {verificationEveryTime}")
            ],
            "Review and execute one-cell mutations on a 1,000-cell workbook and a 100,990-cell, ten-sheet workbook with tables and defined names.",
            ["inspected_cell_count is the exact returned affected-range size, not a COM-call count; the corpus size is recorded separately because the service exposes no COM counter."]);
    }

    private static async Task<StaleResult> ExerciseStaleRangeReviewAsync(
        ExcelMcpService service,
        string sessionId,
        CancellationToken cancellationToken)
    {
        const string args = "{\"sheetName\":\"Data\",\"rangeAddress\":\"C2\",\"values\":[[777]]}";
        var review = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = args,
            ReviewOnly = true,
            Source = "benchmark"
        });
        ServiceBenchmarkHelpers.EnsureSuccess(review, "stale range review");
        var reviewId = BenchmarkContext.GetRequiredString(review.Result, "reviewId");

        var batch = service.SessionManager.GetSession(sessionId)
            ?? throw new InvalidOperationException("Session disappeared during stale-range benchmark.");
        batch.Execute((excelContext, _) =>
        {
            WriteDirectCell(excelContext, "C2", 666d);
            return 0;
        }, cancellationToken);

        var started = Stopwatch.GetTimestamp();
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = args,
            ReviewId = reviewId,
            Source = "benchmark"
        });
        return new StaleResult(
            !response.Success && string.Equals(response.ErrorCategory, "ReviewStale", StringComparison.Ordinal),
            BenchmarkContext.ElapsedMilliseconds(started),
            response.ErrorCategory,
            response.Success ? "Stale review unexpectedly executed." : null);
    }

    private static async Task<StaleResult> ExerciseStaleStructureReviewAsync(
        ExcelMcpService service,
        string sessionId,
        CancellationToken cancellationToken)
    {
        const string args = "{\"sheetName\":\"ReviewedSheet\"}";
        var review = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.create",
            SessionId = sessionId,
            Args = args,
            ReviewOnly = true,
            Source = "benchmark"
        });
        ServiceBenchmarkHelpers.EnsureSuccess(review, "stale structure review");
        var reviewId = BenchmarkContext.GetRequiredString(review.Result, "reviewId");
        var batch = service.SessionManager.GetSession(sessionId)
            ?? throw new InvalidOperationException("Session disappeared during stale-structure benchmark.");
        batch.Execute((excelContext, _) =>
        {
            dynamic? sheets = null;
            dynamic? sheet = null;
            try
            {
                sheets = excelContext.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = $"Intervening{Guid.NewGuid():N}"[..20];
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        }, cancellationToken);

        var started = Stopwatch.GetTimestamp();
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "sheet.create",
            SessionId = sessionId,
            Args = args,
            ReviewId = reviewId,
            Source = "benchmark"
        });
        return new StaleResult(
            !response.Success && string.Equals(response.ErrorCategory, "ReviewStale", StringComparison.Ordinal),
            BenchmarkContext.ElapsedMilliseconds(started),
            response.ErrorCategory,
            response.Success ? "Structurally stale review unexpectedly executed." : null);
    }

    private static void AddStructuralComplexity(string workbookPath)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((excelContext, _) =>
        {
            dynamic? sheets = null;
            dynamic? names = null;
            try
            {
                sheets = excelContext.Book.Worksheets;
                for (var sheetIndex = 2; sheetIndex <= 10; sheetIndex++)
                {
                    dynamic? sheet = null;
                    dynamic? range = null;
                    dynamic? tables = null;
                    dynamic? table = null;
                    try
                    {
                        sheet = sheets.Add();
                        sheet.Name = $"Scope{sheetIndex}";
                        var values = (object[,])Array.CreateInstance(typeof(object), [11, 10], [1, 1]);
                        for (var row = 1; row <= 11; row++)
                        {
                            for (var column = 1; column <= 10; column++)
                            {
                                values[row, column] = row == 1 ? $"H{column}" : row * 100d + column;
                            }
                        }

                        range = sheet.Range["A1:J11"];
                        range.Value2 = values;
                        tables = sheet.ListObjects;
                        table = tables.Add(1, range, Type.Missing, 1, Type.Missing);
                        table.Name = $"ScopeTable{sheetIndex}";
                    }
                    finally
                    {
                        ComUtilities.Release(ref table);
                        ComUtilities.Release(ref tables);
                        ComUtilities.Release(ref range);
                        ComUtilities.Release(ref sheet);
                    }
                }

                names = excelContext.Book.Names;
                for (var nameIndex = 1; nameIndex <= 20; nameIndex++)
                {
                    dynamic? name = null;
                    try
                    {
                        name = names.Add($"BenchmarkName{nameIndex}", $"=Data!$A${nameIndex + 1}");
                    }
                    finally
                    {
                        ComUtilities.Release(ref name);
                    }
                }

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref sheets);
            }
        });
        batch.Save();
    }

    private static void WriteDirectCell(ExcelContext context, string address, double value)
    {
        dynamic? sheet = null;
        dynamic? cell = null;
        try
        {
            sheet = context.Book.Worksheets["Data"];
            cell = sheet.Range[address];
            cell.Value2 = value;
        }
        finally
        {
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref sheet);
        }
    }

    private static bool ReviewHasExactTargetScope(string? result, string sheetName, string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(result))
        {
            return false;
        }

        using var document = JsonDocument.Parse(result);
        if (!document.RootElement.TryGetProperty("affected", out var affected))
        {
            return false;
        }

        var sheets = affected.GetProperty("sheets").EnumerateArray()
            .Select(item => item.GetString())
            .ToArray();
        var ranges = affected.GetProperty("ranges").EnumerateArray()
            .Select(item => item.GetString()?.Replace("$", string.Empty, StringComparison.Ordinal))
            .ToArray();
        var objects = affected.GetProperty("objects").EnumerateArray().ToArray();
        return sheets.Length == 1 &&
            string.Equals(sheets[0], sheetName, StringComparison.OrdinalIgnoreCase) &&
            ranges.Length == 1 &&
            string.Equals(ranges[0], $"{sheetName}!{rangeAddress}", StringComparison.OrdinalIgnoreCase) &&
            objects.Length == 0;
    }

    private static bool HasVerificationStatus(string? result, string expected)
    {
        if (string.IsNullOrWhiteSpace(result))
        {
            return false;
        }

        using var document = JsonDocument.Parse(result);
        return document.RootElement.TryGetProperty("verification", out var verification) &&
            verification.TryGetProperty("status", out var status) &&
            string.Equals(status.GetString(), expected, StringComparison.Ordinal);
    }

    private sealed record StaleResult(bool Rejected, double LatencyMilliseconds, string? Category, string? Error);
}
