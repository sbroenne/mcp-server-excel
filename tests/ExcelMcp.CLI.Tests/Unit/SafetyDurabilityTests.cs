using Sbroenne.ExcelMcp.Service.Safety;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "Safety")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SafetyDurabilityTests : IDisposable
{
    private readonly string _stateRoot = Path.Combine(
        Path.GetTempPath(),
        $"excelmcp-safety-durability-{Guid.NewGuid():N}");

    [Fact]
    public void TryRecordEvidence_WhenStoreWriteThrows_PreservesThePrimaryControlFlow()
    {
        var continued = false;

        var recorded = WorkbookSafetyCoordinator.TryRecordEvidence(
            () => throw new IOException("synthetic journal failure"));
        continued = true;

        Assert.False(recorded);
        Assert.True(continued);
    }

    [Fact]
    public void Transition_WhenSaveFails_RollsBackAndCanRetry()
    {
        var store = new DurableSafetyStore(_stateRoot);
        const string operationId = "retryable-transition";
        store.EnsureOperation(operationId, "session", "range.set-values", "values", "workbook", SafetyScope.Workbook, DateTime.UtcNow);

        BreakJournalDirectory();
        Assert.ThrowsAny<IOException>(() => store.Transition(operationId, "started"));
        Assert.Empty(store.GetJournal("session").Single().Transitions);

        File.Delete(Path.Combine(_stateRoot, "journal"));
        Directory.CreateDirectory(Path.Combine(_stateRoot, "journal"));
        store.Transition(operationId, "started");

        var restarted = new DurableSafetyStore(_stateRoot);
        Assert.Equal("started", restarted.GetJournal("session").Single().Transitions[^1].State);
    }

    [Fact]
    public void TransitionIncomplete_IsolatesWriteFailuresAndRetriesAllOperations()
    {
        var store = new DurableSafetyStore(_stateRoot);
        foreach (var operationId in new[] { "shutdown-one", "shutdown-two" })
        {
            store.EnsureOperation(operationId, "session", "range.set-values", "values", "workbook", SafetyScope.Workbook, DateTime.UtcNow);
        }

        BreakJournalDirectory();
        Assert.Equal(0, store.TransitionIncompleteForSession("session", "abortedUnknown", "ServerShutdown"));

        File.Delete(Path.Combine(_stateRoot, "journal"));
        Directory.CreateDirectory(Path.Combine(_stateRoot, "journal"));
        Assert.Equal(2, store.TransitionIncompleteForSession("session", "abortedUnknown", "ServerShutdown"));
    }

    private void BreakJournalDirectory()
    {
        var journalDirectory = Path.Combine(_stateRoot, "journal");
        Directory.Delete(journalDirectory, recursive: true);
        File.WriteAllText(journalDirectory, "not a directory");
    }

    public void Dispose()
    {
        if (Directory.Exists(_stateRoot))
        {
            Directory.Delete(_stateRoot, recursive: true);
        }

        GC.SuppressFinalize(this);
    }
}
