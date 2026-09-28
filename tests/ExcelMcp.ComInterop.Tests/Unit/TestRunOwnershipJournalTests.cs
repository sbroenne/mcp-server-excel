using System.Text.Json;
using Microsoft.Win32.SafeHandles;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.ComInterop.Tests.Integration;
using Sbroenne.ExcelMcp.Tests.Shared;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionManager")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class TestRunOwnershipJournalTests : IDisposable
{
    private readonly string _directory = Path.Join(Path.GetTempPath(), $"OwnershipJournal_{Guid.NewGuid():N}");

    public TestRunOwnershipJournalTests() => Directory.CreateDirectory(_directory);

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void AssemblyExitVerdict_UsesCapturedHandleInsteadOfReplacementPid(
        bool originalHandleExited, bool replacementPidAppearsExited)
    {
        using var handle = new SafeProcessHandle(new IntPtr(42), ownsHandle: false);
        var identityProbeCalls = 0;
        var handleProbeCalls = 0;
        using var lifetime = new TestRunExcelLifetime(
            _directory,
            _ =>
            {
                identityProbeCalls++;
                return replacementPidAppearsExited;
            },
            capturedHandle =>
            {
                Assert.Same(handle, capturedHandle);
                handleProbeCalls++;
                return originalHandleExited;
            });

        Assert.Equal(originalHandleExited, lifetime.HasExited(new ExcelProcessIdentity(123, 456), handle));
        Assert.Equal(0, identityProbeCalls);
        Assert.Equal(1, handleProbeCalls);
    }

    [Fact]
    public void ConsecutiveExecutions_RearmDisposedLifetimeWithFreshJournal()
    {
        using var first = TestRunExcelLifetime.GetOrCreateActive(null, _directory);
        Assert.Same(first, TestRunExcelLifetime.GetOrCreateActive(first, _directory));
        first.Dispose();

        using var second = TestRunExcelLifetime.GetOrCreateActive(first, _directory);
        Assert.NotSame(first, second);
        Assert.NotEqual(first.JournalPath, second.JournalPath);
        Assert.Same(second, TestRunExcelLifetime.GetOrCreateActive(second, _directory));
        Assert.Contains("\"Kind\":\"host\"", File.ReadAllText(second.JournalPath), StringComparison.Ordinal);
        second.Dispose();
        Assert.Contains("\"Kind\":\"disposed\"", File.ReadAllText(second.JournalPath), StringComparison.Ordinal);
    }

    [Fact]
    public void TestHostStartup_InitializesExcelLifetimeProtection()
    {
        Assert.NotNull(ModuleInit.Lifetime);
        var records = File.ReadAllLines(ModuleInit.Lifetime.JournalPath)
            .Select(line => JsonSerializer.Deserialize<TestRunExcelLifetime.OwnershipRecord>(line));
        var host = Assert.Single(records, record => record?.Kind == "host");
        Assert.NotNull(host);
        Assert.Equal(Environment.ProcessId, host.Host.ProcessId);
        using var process = System.Diagnostics.Process.GetCurrentProcess();
        Assert.Equal(process.StartTime.ToUniversalTime().ToFileTimeUtc(), host.Host.StartedAtUtcFileTime);
    }

    [Fact]
    public void Ready_RequiresSuccessfullyProtectedExcel()
    {
        Write("host-first.jsonl", Line("excel", 101), Line("assignment-failed", 102));
        Assert.Equal(101, TestRunLifetimeTests.ReadOwnedRecord(_directory)?.Excel?.ProcessId);
        Assert.Equal([101, 102], Read().Records.Select(record => record.Excel!.Value.ProcessId));
        Write("host-first.jsonl", Line("assignment-failed", 102));
        Assert.Null(TestRunLifetimeTests.ReadOwnedRecord(_directory));
    }

    [Fact]
    public void Cleanup_MalformedTailPreservesEarlierAndLaterFileIdentities()
    {
        Write("host-first.jsonl", Line("excel", 101), Line("assignment-failed", 102), "{\"Kind\":");
        Write("host-second.jsonl", Line("excel", 103), Line("non-excel-test-process", 104),
            Line("identity-replaced", 105), Line("already-exited", 106));

        var result = Read();

        Assert.Equal([101, 102, 103], result.Records.Select(record => record.Excel!.Value.ProcessId).Order());
        var failure = Assert.IsType<AggregateException>(result.Failure);
        Assert.Single(failure.InnerExceptions);
    }

    [Fact]
    public void Cleanup_MalformedMiddleDoesNotDiscardLaterCompleteRecords()
    {
        Write("host-first.jsonl", Line("excel", 101), "{", Line("assignment-failed", 102));

        var result = Read();

        Assert.Equal([101, 102], result.Records.Select(record => record.Excel!.Value.ProcessId));
        Assert.Single(Assert.IsType<AggregateException>(result.Failure).InnerExceptions);
    }

    [Fact]
    public void Cleanup_ReadErrorDoesNotDiscardReadableFiles()
    {
        Write("host-first.jsonl", Line("excel", 101));
        Write("host-last.jsonl", Line("assignment-failed", 102));
        using var locked = new FileStream(Path.Join(_directory, "host-locked.jsonl"),
            FileMode.Create, FileAccess.ReadWrite, FileShare.None);

        var result = Read();

        Assert.Equal([101, 102], result.Records.Select(record => record.Excel!.Value.ProcessId).Order());
        Assert.Single(Assert.IsType<AggregateException>(result.Failure).InnerExceptions);
    }

    [Fact]
    public void Cleanup_InvalidIdentityRecordsAreReportedAndNeverReturned()
    {
        Write("host-first.jsonl", "null", Line("excel", 0), Line("excel", -1), Line("excel", 101));

        var result = Read();

        Assert.Equal(101, Assert.Single(result.Records).Excel!.Value.ProcessId);
        Assert.Equal(3, Assert.IsType<AggregateException>(result.Failure).InnerExceptions.Count);
    }

    private (List<TestRunExcelLifetime.OwnershipRecord> Records, Exception? Failure) Read()
    {
        var records = new List<TestRunExcelLifetime.OwnershipRecord>();
        var failure = Record.Exception(() =>
        {
            foreach (var record in TestRunLifetimeTests.ReadOwnedRecords(_directory))
            {
                records.Add(record);
            }
        });
        return (records, failure);
    }

    private void Write(string name, params string[] lines) =>
        File.WriteAllLines(Path.Join(_directory, name), lines);

    private static string Line(string kind, int processId) =>
        JsonSerializer.Serialize(new TestRunExcelLifetime.OwnershipRecord(
            kind, new ExcelProcessIdentity(99, 1000), new ExcelProcessIdentity(processId, 2000)));

    public void Dispose() => Directory.Delete(_directory, recursive: true);
}
