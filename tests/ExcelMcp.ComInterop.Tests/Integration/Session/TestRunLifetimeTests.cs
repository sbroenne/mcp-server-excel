using System.Diagnostics;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Shared;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

[Collection("Sequential")]
[Trait("Category", "Integration")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "true")]
public sealed class TestRunLifetimeTests(ITestOutputHelper output)
{
    private const string ChildDirectoryVariable = "EXCELMCP_LIFETIME_PROBE_DIRECTORY";

    [Fact]
    public Task NormalDisposal_ExitsOwnedExcelAndPreservesControl() => RunScenarioAsync("normal");

    [Fact]
    public Task TesthostTermination_ExitsOwnedExcelAndPreservesControl() => RunScenarioAsync("testhost");

    [Fact]
    public Task WrapperTermination_ExitsOwnedExcelAndPreservesControl() => RunScenarioAsync("wrapper");

    [Fact]
    public Task WorkbookWorkflow_UsesBoundedOwnedCleanup() => RunScenarioAsync("workflow");

    [Fact]
    public Task UndisposedExcel_MakesChildHostFailAndPreservesControl() => RunScenarioAsync("undisposed");

    private async Task RunScenarioAsync(string mode, [CallerMemberName] string testName = "")
    {
        var childDirectory = Environment.GetEnvironmentVariable(ChildDirectoryVariable);
        if (childDirectory is not null)
        {
            await RunChildAsync(childDirectory, mode);
            return;
        }

        var directory = Path.Join(Path.GetTempPath(), $"ExcelLifetime_{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        Process? wrapper = null;
        Task<string>? stdout = null;
        Task<string>? stderr = null;
        ExcelProcessIdentity? ownedIdentity = null;
        ExcelProcessIdentity? controlIdentity = null;
        ExcelBatch? control = null;
        Exception? scenarioFailure = null;
        var cleanupFailures = new List<Exception>();
        try
        {
            control = ExcelBatch.CreateNewWorkbook(
                Path.Join(directory, "independent.xlsx"), isMacroEnabled: false);
            Assert.NotNull(control.ExcelProcessId);
            using (var process = Process.GetProcessById(control.ExcelProcessId.Value))
            {
                controlIdentity = new ExcelProcessIdentity(
                    process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());
            }

            var start = new ProcessStartInfo("dotnet")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            start.Environment[ChildDirectoryVariable] = directory;
            start.Environment[TestRunExcelLifetime.OwnershipDirectoryVariable] = directory;
            start.Environment["EXCELMCP_DIAGNOSTICS"] = "1";
            foreach (var argument in new[]
            {
                "vstest", typeof(TestRunLifetimeTests).Assembly.Location,
                $"/TestCaseFilter:FullyQualifiedName={typeof(TestRunLifetimeTests).FullName}.{testName}",
                $"/ResultsDirectory:{directory}",
                "/Logger:trx;LogFileName=child.trx",
                "/Logger:console;verbosity=detailed"
            })
            {
                start.ArgumentList.Add(argument);
            }

            wrapper = Process.Start(start) ?? throw new InvalidOperationException("Child wrapper did not start.");
            stdout = wrapper.StandardOutput.ReadToEndAsync();
            stderr = wrapper.StandardError.ReadToEndAsync();
            var ownershipDirectory = Environment.GetEnvironmentVariable(
                TestRunExcelLifetime.OwnershipDirectoryVariable);
            if (!string.IsNullOrEmpty(ownershipDirectory))
            {
                Directory.CreateDirectory(ownershipDirectory);
                await File.WriteAllTextAsync(Path.Join(ownershipDirectory, "active-probe.json"),
                    JsonSerializer.Serialize(new
                    {
                        Directory = directory,
                        Wrapper = new ExcelProcessIdentity(
                            wrapper.Id, wrapper.StartTime.ToUniversalTime().ToFileTimeUtc())
                    }));
            }
            await WaitUntilAsync(
                () => File.Exists(Path.Join(directory, "ready.json")) || wrapper.HasExited,
                TimeSpan.FromSeconds(60));
            Assert.False(wrapper.HasExited, "Child exited before the ownership READY handshake.");

            var ready = JsonSerializer.Deserialize<TestRunExcelLifetime.OwnershipRecord>(
                await File.ReadAllTextAsync(Path.Join(directory, "ready.json")));
            Assert.NotNull(ready);
            Assert.NotNull(ready.Excel);
            ownedIdentity = ready.Excel;
            Assert.NotEqual(controlIdentity, ownedIdentity);
            var journal = ReadOwnedRecord(directory);
            Assert.Equal(ready, journal);
            output.WriteLine($"READY: host={ready.Host}, owned={ownedIdentity}, control={controlIdentity}");

            if (mode is "normal" or "workflow" or "undisposed")
            {
                await File.WriteAllTextAsync(Path.Join(directory, "stop"), "dispose");
            }
            else if (mode == "testhost")
            {
                Assert.True(OwnedProcessGuard.TryTerminate(
                    ready.Host, TimeSpan.Zero, TimeSpan.FromSeconds(5), out var terminated));
                Assert.True(terminated);
            }
            else
            {
                wrapper.Kill(entireProcessTree: true);
            }

            var exitBudget = mode is "normal" or "workflow" or "undisposed"
                ? ComInteropConstants.StaThreadJoinTimeout
                    + ProcessTerminationPolicy.ProcessExitTimeout * 2
                    + TimeSpan.FromSeconds(15)
                : TimeSpan.FromSeconds(20);
            await wrapper.WaitForExitAsync().WaitAsync(exitBudget);
            if (mode is "normal" or "workflow")
            {
                Assert.Equal(0, wrapper.ExitCode);
            }
            else if (mode == "undisposed")
            {
                Assert.NotEqual(0, wrapper.ExitCode);
                Assert.Contains(
                    Directory.EnumerateFiles(directory, "host-*.jsonl").SelectMany(File.ReadAllLines),
                    line =>
                    {
                        var record = JsonSerializer.Deserialize<TestRunExcelLifetime.OwnershipRecord>(line);
                        return record?.Kind == "normal-teardown-failed" && record.Excel == ownedIdentity;
                    });
                var childOutput = await stdout.WaitAsync(TimeSpan.FromSeconds(10))
                    + await stderr.WaitAsync(TimeSpan.FromSeconds(10));
                Assert.Contains("Owned Excel survived normal testhost teardown", childOutput, StringComparison.Ordinal);
            }

            Assert.True(SpinWait.SpinUntil(
                () => OwnedProcessGuard.TryConfirmExited(ownedIdentity.Value),
                TimeSpan.FromSeconds(10)), $"Owned Excel survived {mode}: {ownedIdentity}.");
            Assert.True(control.IsExcelProcessAlive());
            control.Execute((context, _) => Assert.Equal("independent.xlsx", context.Book.Name));
        }
        catch (Exception ex)
        {
            scenarioFailure = ex;
        }
        finally
        {
            try
            {
                if (wrapper is not null)
                {
                    if (!wrapper.HasExited)
                    {
                        wrapper.Kill(entireProcessTree: true);
                    }

                    await wrapper.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
                    if (stdout is not null && stderr is not null)
                    {
                        output.WriteLine(await stdout.WaitAsync(TimeSpan.FromSeconds(10)));
                        output.WriteLine(await stderr.WaitAsync(TimeSpan.FromSeconds(10)));
                    }
                }
            }
            catch (Exception ex)
            {
                cleanupFailures.Add(ex);
            }
            finally
            {
                CaptureCleanup(() => wrapper?.Dispose());
                var recordedIdentities = new HashSet<ExcelProcessIdentity>();
                if (ownedIdentity is { } readyIdentity) recordedIdentities.Add(readyIdentity);
                CaptureCleanup(() =>
                {
                    foreach (var record in ReadOwnedRecords(directory))
                    {
                        if (record.Excel is { } identity) recordedIdentities.Add(identity);
                    }
                });
                foreach (var identity in recordedIdentities)
                {
                    CaptureCleanup(() =>
                    {
                        Assert.True(OwnedProcessGuard.TryTerminate(
                            identity, TimeSpan.Zero, TimeSpan.FromSeconds(5), out _),
                            $"Failed to clean the recorded child Excel identity {identity}.");
                    });
                }
                CaptureCleanup(() => control?.Dispose());
                CaptureCleanup(() =>
                {
                    if (controlIdentity is { } independent)
                    {
                        Assert.True(OwnedProcessGuard.TryConfirmExited(independent),
                            $"Control Excel did not exit through normal disposal: {independent}.");
                    }
                });
            }

            CaptureCleanup(() =>
            {
                var evidenceDirectory = Environment.GetEnvironmentVariable(
                    TestRunExcelLifetime.OwnershipDirectoryVariable);
                if (!string.IsNullOrEmpty(evidenceDirectory))
                {
                    var destination = Path.Join(evidenceDirectory, $"probe-{testName}-{Guid.NewGuid():N}");
                    Directory.CreateDirectory(destination);
                    foreach (var file in Directory.EnumerateFiles(directory))
                    {
                        if (!Path.GetExtension(file).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            File.Copy(file, Path.Join(destination, Path.GetFileName(file)));
                        }
                    }
                }
                Directory.Delete(directory, recursive: true);
            });
        }

        if (cleanupFailures.Count > 0)
        {
            if (scenarioFailure is not null) cleanupFailures.Insert(0, scenarioFailure);
            throw new AggregateException("Lifecycle scenario or cleanup failed.", cleanupFailures);
        }
        if (scenarioFailure is not null) ExceptionDispatchInfo.Capture(scenarioFailure).Throw();

        void CaptureCleanup(Action cleanup)
        {
            try { cleanup(); }
            catch (Exception ex) { cleanupFailures.Add(ex); }
        }
    }

    private static async Task RunChildAsync(string directory, string mode)
    {
        Assert.NotNull(ModuleInit.Lifetime);
        var logger = new ShutdownLogger(Path.Join(directory, "shutdown.log"));
        var batch = ExcelBatch.CreateNewWorkbook(
            Path.Join(directory, "owned.xlsx"), isMacroEnabled: false, logger);
        try
        {
            batch.Execute((context, _) =>
            {
                var security = context.App.GetType().InvokeMember(
                    "AutomationSecurity", System.Reflection.BindingFlags.GetProperty,
                    binder: null, target: context.App, args: null, culture: CultureInfo.InvariantCulture);
                Assert.Equal(3, Convert.ToInt32(security, CultureInfo.InvariantCulture));
            });
            if (mode == "workflow")
            {
                ExerciseWorksheetTable(batch, create: true);
                batch.Save();
            }
            var record = ReadOwnedRecord(directory);
            Assert.NotNull(record);
            Assert.NotNull(record.Excel);
            await File.WriteAllTextAsync(Path.Join(directory, "ready.tmp"), JsonSerializer.Serialize(record));
            File.Move(Path.Join(directory, "ready.tmp"), Path.Join(directory, "ready.json"));
            await WaitUntilAsync(() => File.Exists(Path.Join(directory, "stop")), TimeSpan.FromSeconds(90));
            if (mode == "undisposed")
            {
                // Deliberately let the test pass with a live session. Host teardown must
                // turn this into a failing runner exit, not just a warning beside green TRX.
                return;
            }
            batch.Dispose();
            await AssertBoundedCleanupAsync(logger, 0, record.Excel.Value,
                Path.Join(directory, "initial-cleanup.json"), requireSave: mode == "workflow");
            if (mode == "workflow")
            {
                using var reopened = new ExcelBatch([Path.Join(directory, "owned.xlsx")], logger);
                var messageCount = logger.Messages.Length;
                ExerciseWorksheetTable(reopened, create: false);
                reopened.Save();
                var reopenedRecord = ReadOwnedRecord(directory);
                Assert.NotNull(reopenedRecord?.Excel);
                var reopenedIdentity = reopenedRecord.Excel.Value;
                Assert.NotEqual(record.Excel.Value, reopenedIdentity);
                reopened.Dispose();
                await AssertBoundedCleanupAsync(logger, messageCount, reopenedIdentity,
                    Path.Join(directory, "reopen-cleanup.json"), requireSave: true);
            }
            await File.WriteAllTextAsync(Path.Join(directory, "disposed"), "normal cleanup confirmed");
        }
        finally
        {
            if (mode != "undisposed") batch.Dispose();
        }
    }

    private static async Task AssertBoundedCleanupAsync(
        ShutdownLogger logger, int messageStart, ExcelProcessIdentity identity,
        string evidencePath, bool requireSave)
    {
        var messages = logger.Messages.Skip(messageStart).ToArray();
        if (requireSave) Assert.Contains("Workbook owned.xlsx saved successfully", messages);
        Assert.Contains("Workbook owned.xlsx closed successfully", messages);
        Assert.Contains(messages, message => message.Contains("[DIAG-QUIT-SUCCESS]", StringComparison.Ordinal));
        Assert.Contains("STA thread cleanup completed for owned.xlsx", messages);
        Assert.DoesNotContain(messages, message =>
            message.Contains("[DIAG-DISPOSE-STA-JOIN-TIMEOUT]", StringComparison.Ordinal)
            || message.Contains("[DIAG-DISPOSE-FORCE-KILL", StringComparison.Ordinal)
            || message.Contains("[DIAG-DISPOSE-TIMEOUT-PREKILL", StringComparison.Ordinal)
            || (message.Contains("[DIAG-QUIT-", StringComparison.Ordinal)
                && !message.Contains("[DIAG-QUIT-SUCCESS]", StringComparison.Ordinal))
            || message.Contains("Failed to close workbook", StringComparison.Ordinal)
            || message.Contains("COM proxy disconnected", StringComparison.Ordinal));
        Assert.NotNull(ModuleInit.Lifetime);
        Assert.True(ModuleInit.Lifetime.HasExited(identity),
            "The captured Excel process must exit before the testhost job is disposed.");
        Assert.NotNull(logger.LastProcessWaitElapsed);
        Assert.InRange(logger.LastProcessWaitElapsed.Value,
            TimeSpan.Zero, ProcessTerminationPolicy.NormalShutdownBudget);
        var fallback = messages
            .Where(message => message.Contains("[DIAG-DISPOSE-PROCESS-LINGER-KILLED]", StringComparison.Ordinal))
            .ToArray();
        if (fallback.Length > 0)
        {
            Assert.Contains($"Excel process {identity.ProcessId} ", Assert.Single(fallback), StringComparison.Ordinal);
        }
        await File.WriteAllTextAsync(evidencePath, JsonSerializer.Serialize(new
        {
            identity,
            usedPostQuitProcessFallback = fallback.Length > 0,
            postQuitProcessWaitSeconds = logger.LastProcessWaitElapsed.Value.TotalSeconds,
            processWaitBudgetSeconds = ProcessTerminationPolicy.NormalShutdownBudget.TotalSeconds,
            exitedBeforeJobDisposal = true
        }));
    }

    private static void ExerciseWorksheetTable(ExcelBatch batch, bool create) =>
        batch.Execute((context, _) =>
        {
            const string tableName = "\u58F2\u4E0A_\u5B9F\u7E3E";
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                range = sheet.Range["A1:B2"];
                tables = sheet.ListObjects;
                if (create)
                {
                    range.Value2 = new object[,] { { "Label", "Amount" }, { "example", 42d } };
                    table = tables.Add(Excel.XlListObjectSourceType.xlSrcRange, range,
                        Type.Missing, Excel.XlYesNoGuess.xlYes);
                    ((dynamic)(object)table).Name = tableName;
                }
                else
                {
                    table = tables[tableName];
                }

                Assert.Equal(tableName, table.Name);
                var values = (object[,])range.Value2;
                Assert.Equal(42d, values[2, 2]);
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    internal static TestRunExcelLifetime.OwnershipRecord? ReadOwnedRecord(string directory) =>
        ReadOwnedRecords(directory).LastOrDefault(record => record.Kind == "excel");

    internal static IEnumerable<TestRunExcelLifetime.OwnershipRecord> ReadOwnedRecords(string directory)
    {
        var records = new List<TestRunExcelLifetime.OwnershipRecord>();
        var failures = new List<Exception>();
        try
        {
            foreach (var path in Directory.EnumerateFiles(directory, "host-*.jsonl"))
            {
                try
                {
                    using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var reader = new StreamReader(stream);
                    var lineNumber = 0;
                    while (reader.ReadLine() is { } line)
                    {
                        lineNumber++;
                        try
                        {
                            var record = JsonSerializer.Deserialize<TestRunExcelLifetime.OwnershipRecord>(line);
                            if (record is null || string.IsNullOrEmpty(record.Kind))
                            {
                                throw new InvalidDataException("Missing ownership record kind.");
                            }
                            if (record.Kind is not ("excel" or "assignment-failed")) continue;
                            if (record.Host.ProcessId <= 0 || record.Host.StartedAtUtcFileTime <= 0
                                || record.Excel is not { ProcessId: > 0, StartedAtUtcFileTime: > 0 })
                            {
                                throw new InvalidDataException("Missing valid host or Excel process identity.");
                            }
                            records.Add(record);
                        }
                        catch (Exception ex) when (ex is JsonException or InvalidDataException)
                        {
                            failures.Add(new InvalidDataException(
                                $"Invalid ownership record in {Path.GetFileName(path)} at line {lineNumber}.", ex));
                        }
                    }
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    failures.Add(new IOException($"Cannot finish reading ownership journal {Path.GetFileName(path)}.", ex));
                }
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            failures.Add(new IOException("Cannot finish enumerating ownership journals.", ex));
        }

        // Deliver every usable identity before reporting errors so cleanup cannot lose
        // earlier records, later records, or another journal because one read failed.
        foreach (var record in records) yield return record;
        if (failures.Count > 0) throw new AggregateException("Ownership journal reading failed.", failures);
    }

    private static async Task WaitUntilAsync(Func<bool> condition, TimeSpan timeout)
    {
        using var deadline = new CancellationTokenSource(timeout);
        while (!condition())
        {
            await Task.Delay(50, deadline.Token);
        }
    }

    private sealed class ShutdownLogger(string path) : ILogger<ExcelBatch>
    {
        private readonly object _gate = new();
        private readonly List<string> _messages = [];
        private long _processWaitStarted;

        internal TimeSpan? LastProcessWaitElapsed { get; private set; }

        internal string[] Messages
        {
            get { lock (_gate) { return _messages.ToArray(); } }
        }

        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;

        public bool IsEnabled(LogLevel logLevel) => true;

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
            Exception? exception, Func<TState, Exception?, string> formatter)
        {
            var message = formatter(state, exception);
            lock (_gate)
            {
                if (message.Contains("[DIAG-DISPOSE-PROCESS-WAIT]", StringComparison.Ordinal))
                {
                    _processWaitStarted = Stopwatch.GetTimestamp();
                }
                else if (message.Contains("Dispose COMPLETED", StringComparison.Ordinal) && _processWaitStarted != 0)
                {
                    LastProcessWaitElapsed = Stopwatch.GetElapsedTime(_processWaitStarted);
                    _processWaitStarted = 0;
                }
                _messages.Add(message);
                File.AppendAllText(path,
                    $"{DateTime.UtcNow:o} thread={Environment.CurrentManagedThreadId} {message}{Environment.NewLine}");
            }
        }
    }
}
