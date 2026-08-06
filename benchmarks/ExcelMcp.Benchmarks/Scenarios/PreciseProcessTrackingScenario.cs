using System.Diagnostics;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Benchmarks.Scenarios;

internal sealed class PreciseProcessTrackingScenario : IBenchmarkScenario
{
    public string PlanId => "09";

    public string Name => "precise-process-tracking";

    public Task<ScenarioResult> RunAsync(BenchmarkContext context, CancellationToken cancellationToken)
    {
        var targetMaster = context.CreateWorkingPath("process-target-master");
        var sentinelMaster = context.CreateWorkingPath("process-sentinel-master");
        BenchmarkContext.CreateDataWorkbook(targetMaster, rows: 2, columns: 1);
        BenchmarkContext.CreateDataWorkbook(sentinelMaster, rows: 2, columns: 1);
        var observations = new List<BenchmarkObservation>();
        var sentinelAlwaysSurvived = true;
        var ownedAlwaysExited = true;
        var identityMismatchAlwaysFailedClosed = true;
        var identityValidationAlwaysSupported = true;
        var wrongProcessKills = 0;
        var total = context.Configuration.Warmups + context.Configuration.ReliabilityIterations;

        for (var iteration = 0; iteration < total; iteration++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var targetPath = context.CopyWorkbook(targetMaster, "process-target");
            var sentinelPath = context.CopyWorkbook(sentinelMaster, "process-sentinel");
            IExcelBatch? sentinel = null;
            SessionManager? manager = null;
            int? sentinelPid = null;
            int? targetPid = null;
            DateTime? sentinelStart = null;
            DateTime? targetStart = null;
            var targetExited = false;
            var sentinelSurvived = false;
            var cleanupLatency = 0d;
            var identityMismatchDetectionLatency = 0d;
            var identityMismatchFailedClosed = false;
            var identityValidationSupported = false;
            var error = default(string);

            try
            {
                sentinel = ExcelSession.BeginBatch(context.Configuration.ShowExcel, operationTimeout: null, sentinelPath);
                sentinelPid = sentinel.ExcelProcessId;
                sentinelStart = ReadProcessStartTime(sentinelPid);
                manager = new SessionManager();
                var sessionId = manager.CreateSession(targetPath, context.Configuration.ShowExcel);
                var target = manager.GetSession(sessionId)
                    ?? throw new InvalidOperationException("Target process session was not registered.");
                targetPid = target.ExcelProcessId;
                targetStart = ReadProcessStartTime(targetPid);

                var identity = (target as ExcelBatch)?.ExcelProcessIdentity;
                identityValidationSupported = identity.HasValue;
                if (identity.HasValue)
                {
                    var staleIdentity = identity.Value with
                    {
                        StartedAtUtcFileTime = identity.Value.StartedAtUtcFileTime - 1
                    };
                    var mismatchStarted = Stopwatch.GetTimestamp();
                    var staleIdentityAlive = OwnedProcessIdentityGuard.IsAlive(staleIdentity);
                    var staleIdentityKilled = OwnedProcessIdentityGuard.TryKill(staleIdentity);
                    identityMismatchDetectionLatency = BenchmarkContext.ElapsedMilliseconds(mismatchStarted);
                    identityMismatchFailedClosed =
                        !staleIdentityAlive &&
                        !staleIdentityKilled &&
                        targetPid.HasValue &&
                        BenchmarkContext.IsProcessAlive(targetPid.Value);
                }

                var cleanupStarted = Stopwatch.GetTimestamp();
                _ = manager.CloseSession(sessionId, save: false, force: false);
                if (targetPid.HasValue)
                {
                    targetExited = BenchmarkContext.WaitForProcessExit(
                        targetPid.Value,
                        TimeSpan.FromSeconds(15),
                        out _);
                }

                cleanupLatency = BenchmarkContext.ElapsedMilliseconds(cleanupStarted);
                sentinelSurvived = sentinelPid.HasValue &&
                    BenchmarkContext.IsProcessAlive(sentinelPid.Value) &&
                    sentinel.IsExcelProcessAlive();
            }
            catch (Exception exception)
            {
                error = $"{exception.GetType().Name}: {exception.Message}";
            }
            finally
            {
                manager?.Dispose();
                sentinel?.Dispose();
            }

            sentinelAlwaysSurvived &= sentinelSurvived;
            ownedAlwaysExited &= targetExited;
            identityMismatchAlwaysFailedClosed &= identityMismatchFailedClosed;
            identityValidationAlwaysSupported &= identityValidationSupported;
            if (!sentinelSurvived)
            {
                wrongProcessKills++;
            }

            if (iteration >= context.Configuration.Warmups)
            {
                observations.Add(new BenchmarkObservation(
                    iteration - context.Configuration.Warmups,
                    "two-owned-excel-processes",
                    error is null && targetExited && sentinelSurvived && identityMismatchFailedClosed,
                    error,
                    new Dictionary<string, double>
                    {
                        ["owned_process_exit_ms"] = cleanupLatency,
                        ["orphan_process_count"] = targetExited ? 0 : 1,
                        ["wrong_process_kill_count"] = sentinelSurvived ? 0 : 1,
                        ["identity_mismatch_detection_ms"] = identityMismatchDetectionLatency,
                        ["cleanup_success_rate"] = targetExited ? 1 : 0,
                        ["identity_validation_supported"] = identityValidationSupported ? 1 : 0
                    },
                    new Dictionary<string, string>
                    {
                        ["target_pid"] = targetPid?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "not-captured",
                        ["target_start_utc"] = targetStart?.ToUniversalTime().ToString("O", System.Globalization.CultureInfo.InvariantCulture) ?? "not-captured",
                        ["sentinel_pid"] = sentinelPid?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "not-captured",
                        ["sentinel_start_utc"] = sentinelStart?.ToUniversalTime().ToString("O", System.Globalization.CultureInfo.InvariantCulture) ?? "not-captured",
                        ["implementation_identity"] = "pid+start-time+process-name+executable-path"
                    },
                    targetExited && sentinelSurvived && identityMismatchFailedClosed ? "isolated-identity-checked-cleanup" : "cleanup-failure"));
            }
        }

        var plan = BenchmarkPlanCatalog.All.Single(plan => plan.Id == PlanId);
        return Task.FromResult(ScenarioResult.Create(
            PlanId,
            Name,
            plan.Title,
            observations,
            [
                new BenchmarkInvariant("sentinel_process_survives", sentinelAlwaysSurvived, $"The separately owned sentinel remained alive during every target close: {sentinelAlwaysSurvived}"),
                new BenchmarkInvariant("identity_mismatch_fails_closed", identityMismatchAlwaysFailedClosed, $"Every deliberately stale process identity was rejected without terminating its live PID: {identityMismatchAlwaysFailedClosed}"),
                new BenchmarkInvariant("no_wrong_process_kill", wrongProcessKills == 0, $"Observed sentinel terminations: {wrongProcessKills}"),
                new BenchmarkInvariant("no_owned_excel_orphan", ownedAlwaysExited, $"Every target Excel process exited within 15 seconds: {ownedAlwaysExited}")
            ],
            "Open two isolated real-Excel processes, close one through SessionManager, verify the exact target exits, and verify the sentinel process remains usable.",
            [
                $"Full identity validation was captured in every iteration: {identityValidationAlwaysSupported}.",
                "Each control path validates PID, process start time, process name, and executable path before probing, waiting for, or killing Excel."
            ]));
    }

    private static DateTime? ReadProcessStartTime(int? processId)
    {
        if (!processId.HasValue)
        {
            return null;
        }

        try
        {
            using var process = Process.GetProcessById(processId.Value);
            return process.StartTime;
        }
        catch (ArgumentException)
        {
            return null;
        }
        catch (InvalidOperationException)
        {
            return null;
        }
    }
}
