using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;

namespace Sbroenne.ExcelMcp.Benchmarks.Scenarios;

internal sealed class DurableJournalCheckpointScenario : IBenchmarkScenario
{
    public string PlanId => "08";

    public string Name => "durable-journal-checkpoints";

    public async Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken)
    {
        var masterPath = context.CreateWorkingPath("durability-master");
        BenchmarkContext.CreateDataWorkbook(masterPath, rows: 100, columns: 10);
        var observations = new List<BenchmarkObservation>();
        var journalParseable = true;
        var transitionOrderValid = true;
        var checkpointHashesValid = true;
        var recoveredStateExact = true;
        var restartFindsEveryCheckpoint = true;
        var noTemporaryArtifacts = true;
        var total = context.Configuration.Warmups + context.Configuration.Iterations;

        for (var iteration = 0; iteration < total; iteration++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var workbookPath = context.CopyWorkbook(masterPath, "durability");
            var safetyRoot = context.CreateSafetyRoot("durability");
            var service = new ExcelMcpService(safetyRoot);
            var sessionId = await ServiceBenchmarkHelpers.CreateSessionAsync(
                service,
                workbookPath,
                context.Configuration.ShowExcel);
            await ServiceBenchmarkHelpers.ConfigureSafetyAsync(
                service,
                sessionId,
                reviewMode: "off",
                checkpointMode: "required");

            var mutationRequest = new ServiceRequest
            {
                Command = "range.set-values",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new
                {
                    sheetName = "Data",
                    rangeAddress = "A2",
                    values = new object?[][] { [700_000d + iteration] }
                }, ServiceProtocol.JsonOptions),
                Source = "benchmark"
            };
            var mutationStarted = Stopwatch.GetTimestamp();
            var mutation = await service.ProcessAsync(mutationRequest);
            var checkpointLatency = BenchmarkContext.ElapsedMilliseconds(mutationStarted);
            ServiceBenchmarkHelpers.EnsureSuccess(mutation, "checkpointed mutation");
            var checkpoint = ReadCheckpoint(mutation.Result);
            var hashValid = File.Exists(checkpoint.Path) &&
                string.Equals(ComputeFileHash(checkpoint.Path), checkpoint.Sha256, StringComparison.OrdinalIgnoreCase) &&
                new FileInfo(checkpoint.Path).Length == checkpoint.Size;
            checkpointHashesValid &= hashValid;

            var journalStarted = Stopwatch.GetTimestamp();
            var journal = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.journal",
                SessionId = sessionId,
                Source = "benchmark"
            });
            var journalLatency = BenchmarkContext.ElapsedMilliseconds(journalStarted);
            ServiceBenchmarkHelpers.EnsureSuccess(journal, "session.journal");
            var transitions = ReadOperationTransitions(journal.Result, checkpoint.OperationId);
            var ordered = ContainsOrderedSubsequence(
                transitions,
                ["checkpointReserved", "checkpointCreated", "started", "completed", "verified"]);
            transitionOrderValid &= ordered;
            var parseable = AllJournalFilesParse(safetyRoot);
            journalParseable &= parseable;
            var temporaryArtifactsAbsent = NoTemporaryArtifacts(safetyRoot);
            noTemporaryArtifacts &= temporaryArtifactsAbsent;

            // Dispose without an explicit close to exercise the abnormal-shutdown recovery policy.
            service.Dispose();

            using var restartedService = new ExcelMcpService(safetyRoot);
            var restartStarted = Stopwatch.GetTimestamp();
            var recoveries = await restartedService.ProcessAsync(new ServiceRequest
            {
                Command = "recovery.list",
                Source = "benchmark"
            });
            var restartLatency = BenchmarkContext.ElapsedMilliseconds(restartStarted);
            ServiceBenchmarkHelpers.EnsureSuccess(recoveries, "recovery.list after restart");
            var recoveryFound = RecoveryListContains(recoveries.Result, checkpoint.RecoveryId);
            restartFindsEveryCheckpoint &= recoveryFound;

            var recoverStarted = Stopwatch.GetTimestamp();
            var recover = await restartedService.ProcessAsync(new ServiceRequest
            {
                Command = "recovery.recover",
                Args = JsonSerializer.Serialize(new
                {
                    recoveryId = checkpoint.RecoveryId,
                    show = context.Configuration.ShowExcel
                }, ServiceProtocol.JsonOptions),
                Source = "benchmark"
            });
            ServiceBenchmarkHelpers.EnsureSuccess(recover, "recovery.recover");
            var recoveredSessionId = BenchmarkContext.GetRequiredString(recover.Result, "sessionId");
            var recoveredValue = await ReadSingleNumberAsync(restartedService, recoveredSessionId, "A2");
            var recoveryLatency = BenchmarkContext.ElapsedMilliseconds(recoverStarted);
            var exact = Math.Abs(recoveredValue - 1001d) < 0.000001;
            recoveredStateExact &= exact;
            await ServiceBenchmarkHelpers.CloseSessionAsync(restartedService, recoveredSessionId);

            if (iteration >= context.Configuration.Warmups)
            {
                observations.Add(new BenchmarkObservation(
                    iteration - context.Configuration.Warmups,
                    "checkpoint-restart-recovery",
                    hashValid && ordered && parseable && temporaryArtifactsAbsent && recoveryFound && exact,
                    null,
                    new Dictionary<string, double>
                    {
                        ["checkpoint_create_ms"] = checkpointLatency,
                        ["journal_write_ms"] = checkpointLatency,
                        ["journal_read_ms"] = journalLatency,
                        ["checkpoint_bytes"] = checkpoint.Size,
                        ["restart_recovery_ms"] = restartLatency,
                        ["refresh_to_consistent_read_ms"] = recoveryLatency
                    },
                    new Dictionary<string, string>
                    {
                        ["checkpoint_sha256"] = checkpoint.Sha256,
                        ["checkpoint_recovery_id"] = checkpoint.RecoveryId,
                        ["crash_simulation"] = "service-dispose-without-session-close",
                        ["journal_publication"] = "flush-to-disk-atomic-replace",
                        ["checkpoint_publication"] = "staged-flush-atomic-move"
                    },
                    "recovered-pre-write-state"));
            }
        }

        var failsClosed = await VerifyRequiredCheckpointFailsClosedAsync(context, masterPath, cancellationToken);
        var corruptJournalFailsClosed = VerifyCorruptJournalFailsClosed(context);
        observations.Add(new BenchmarkObservation(
            context.Configuration.Iterations,
            "required-checkpoint-outage",
            failsClosed,
            failsClosed ? null : "Mutation executed even though the required checkpoint path was unavailable.",
            new Dictionary<string, double> { ["required_checkpoint_blocked"] = failsClosed ? 1 : 0 },
            Outcome: failsClosed ? "blocked" : "unexpected-write"));

        var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == PlanId);
        return ScenarioResult.Create(
            PlanId,
            Name,
            plan.Title,
            observations,
            [
                new BenchmarkInvariant("journal_parseable", journalParseable, $"Every durable JSON record parsed after mutation: {journalParseable}"),
                new BenchmarkInvariant("transition_order", transitionOrderValid, $"Every operation preserved checkpoint→started→completed→verified ordering: {transitionOrderValid}"),
                new BenchmarkInvariant("checkpoint_hash_valid", checkpointHashesValid, $"Every checkpoint size and SHA-256 matched: {checkpointHashesValid}"),
                new BenchmarkInvariant("required_checkpoint_fails_closed", failsClosed, $"Synthetic checkpoint-directory outage blocked the write: {failsClosed}"),
                new BenchmarkInvariant("corrupt_journal_fails_closed", corruptJournalFailsClosed, $"A truncated journal blocked service startup: {corruptJournalFailsClosed}"),
                new BenchmarkInvariant("no_temporary_artifacts", noTemporaryArtifacts, $"Every acknowledged transition left no journal/checkpoint staging file: {noTemporaryArtifacts}"),
                new BenchmarkInvariant("recovered_state_exact", recoveredStateExact && restartFindsEveryCheckpoint, $"Every restart found its checkpoint and recovered original A2=1001: {recoveredStateExact && restartFindsEveryCheckpoint}")
            ],
            "Checkpoint a mutation, verify the checkpoint hash and journal order, dispose without session.close, restart the service, recover, and compare the pre-write cell value.",
            [
                "journal_write_ms currently measures the full journaled+checkpointed mutation because the service exposes no transition-level timer.",
                "Journal and checkpoint bytes are flushed before same-directory atomic publication; process disposal remains a crash-like restart test rather than a literal power-cut test."
            ]);
    }

    private static bool VerifyCorruptJournalFailsClosed(BenchmarkContext context)
    {
        var safetyRoot = context.CreateSafetyRoot("corrupt-journal");
        var journalDirectory = Path.Combine(safetyRoot, "journal");
        Directory.CreateDirectory(journalDirectory);
        File.WriteAllText(Path.Combine(journalDirectory, "truncated.json"), "{\"operationId\":");

        try
        {
            using var _ = new ExcelMcpService(safetyRoot);
            return false;
        }
        catch (InvalidDataException)
        {
            return true;
        }
    }

    private static async Task<bool> VerifyRequiredCheckpointFailsClosedAsync(
        BenchmarkContext context,
        string masterPath,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var workbookPath = context.CopyWorkbook(masterPath, "checkpoint-outage");
        var safetyRoot = context.CreateSafetyRoot("checkpoint-outage");
        using var service = new ExcelMcpService(safetyRoot);
        var sessionId = await ServiceBenchmarkHelpers.CreateSessionAsync(
            service,
            workbookPath,
            context.Configuration.ShowExcel);
        await ServiceBenchmarkHelpers.ConfigureSafetyAsync(
            service,
            sessionId,
            reviewMode: "off",
            checkpointMode: "required");

        var checkpointsPath = Path.Combine(safetyRoot, "checkpoints");
        if (Directory.Exists(checkpointsPath))
        {
            Directory.Delete(checkpointsPath, recursive: true);
        }
        File.WriteAllText(checkpointsPath, "synthetic checkpoint outage");

        var mutation = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.set-values",
            SessionId = sessionId,
            Args = "{\"sheetName\":\"Data\",\"rangeAddress\":\"A2\",\"values\":[[999999]]}",
            Source = "benchmark"
        });
        var value = await ReadSingleNumberAsync(service, sessionId, "A2");
        await ServiceBenchmarkHelpers.CloseSessionAsync(service, sessionId);
        return !mutation.Success && Math.Abs(value - 1001d) < 0.000001;
    }

    private static CheckpointInfo ReadCheckpoint(string? result)
    {
        using var document = JsonDocument.Parse(result ?? throw new InvalidDataException("Checkpoint mutation returned no JSON."));
        var root = document.RootElement;
        var checkpoint = root.GetProperty("checkpoint");
        return new CheckpointInfo(
            root.GetProperty("operationId").GetString() ?? throw new InvalidDataException("Missing operationId."),
            checkpoint.GetProperty("recoveryId").GetString() ?? throw new InvalidDataException("Missing recoveryId."),
            checkpoint.GetProperty("path").GetString() ?? throw new InvalidDataException("Missing checkpoint path."),
            checkpoint.GetProperty("sha256").GetString() ?? throw new InvalidDataException("Missing checkpoint SHA-256."),
            checkpoint.GetProperty("size").GetInt64());
    }

    private static string[] ReadOperationTransitions(string? result, string operationId)
    {
        using var document = JsonDocument.Parse(result ?? throw new InvalidDataException("Journal returned no JSON."));
        foreach (var operation in document.RootElement.GetProperty("operations").EnumerateArray())
        {
            if (string.Equals(operation.GetProperty("operationId").GetString(), operationId, StringComparison.Ordinal))
            {
                return operation.GetProperty("transitions")
                    .EnumerateArray()
                    .Select(item => item.GetProperty("state").GetString() ?? string.Empty)
                    .ToArray();
            }
        }

        return [];
    }

    private static bool ContainsOrderedSubsequence(IReadOnlyList<string> actual, IReadOnlyList<string> expected)
    {
        var expectedIndex = 0;
        foreach (var item in actual)
        {
            if (expectedIndex < expected.Count && string.Equals(item, expected[expectedIndex], StringComparison.Ordinal))
            {
                expectedIndex++;
            }
        }

        return expectedIndex == expected.Count;
    }

    private static bool AllJournalFilesParse(string safetyRoot)
    {
        var journalDirectory = Path.Combine(safetyRoot, "journal");
        var files = Directory.Exists(journalDirectory)
            ? Directory.EnumerateFiles(journalDirectory, "*.json", SearchOption.TopDirectoryOnly).ToArray()
            : [];
        if (files.Length == 0)
        {
            return false;
        }

        try
        {
            foreach (var file in files)
            {
                using var _ = JsonDocument.Parse(File.ReadAllText(file));
            }

            return true;
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static bool NoTemporaryArtifacts(string safetyRoot) =>
        !Directory.EnumerateFiles(safetyRoot, "*.tmp", SearchOption.AllDirectories).Any() &&
        !Directory.EnumerateFiles(safetyRoot, ".*.pending.*", SearchOption.AllDirectories).Any();

    private static bool RecoveryListContains(string? result, string recoveryId)
    {
        using var document = JsonDocument.Parse(result ?? throw new InvalidDataException("Recovery list returned no JSON."));
        return document.RootElement.GetProperty("recoveries")
            .EnumerateArray()
            .Any(item => string.Equals(item.GetProperty("recoveryId").GetString(), recoveryId, StringComparison.Ordinal) &&
                item.GetProperty("available").GetBoolean());
    }

    private static async Task<double> ReadSingleNumberAsync(ExcelMcpService service, string sessionId, string address)
    {
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "range.get-values",
            SessionId = sessionId,
            Args = JsonSerializer.Serialize(new { sheetName = "Data", rangeAddress = address }, ServiceProtocol.JsonOptions),
            Source = "benchmark"
        });
        ServiceBenchmarkHelpers.EnsureSuccess(response, $"read Data!{address}");
        using var document = JsonDocument.Parse(response.Result!);
        return document.RootElement.GetProperty("values")[0][0].GetDouble();
    }

    private static string ComputeFileHash(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private sealed record CheckpointInfo(string OperationId, string RecoveryId, string Path, string Sha256, long Size);
}
