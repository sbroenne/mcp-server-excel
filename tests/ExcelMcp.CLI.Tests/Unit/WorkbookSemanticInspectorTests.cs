using Sbroenne.ExcelMcp.Service.Safety;
using System.Runtime.InteropServices;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "Safety")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class WorkbookSemanticInspectorTests
{
    [Fact]
    public void Compare_ExactRangeEvidence_ReturnsVerified()
    {
        var rangeScope = new SafetyScope(["Sheet1"], ["Sheet1!$A$1"], []);
        var before = Snapshot("before", "rangeSemantic", ["a"], bounded: true, scope: rangeScope);
        var after = Snapshot("after", "rangeSemantic", ["b"], bounded: true, scope: rangeScope);

        var receipt = WorkbookSemanticInspector.Compare(before, after);

        Assert.Equal("verified", receipt.Status);
        Assert.Equal(1, receipt.ChangedCells);
        Assert.Same(before.Scope, receipt.Scope);
        Assert.Null(receipt.Limitation);
    }

    [Fact]
    public void Compare_ExactRangeReceiptUsesComparableTargetFingerprints()
    {
        var rangeScope = new SafetyScope(["Sheet1"], ["Sheet1!$A$1"], []);
        var before = new SemanticSnapshot(
            "authorization-includes-workbook-structure",
            "target-before",
            rangeScope,
            "rangeSemantic",
            true,
            ["a"],
            1);
        var after = new SemanticSnapshot(
            "post-write-target-only",
            "target-after",
            rangeScope,
            "rangeSemantic",
            true,
            ["b"],
            1);

        var receipt = WorkbookSemanticInspector.Compare(before, after);

        Assert.Equal("target-before", receipt.BeforeFingerprint);
        Assert.Equal("target-after", receipt.AfterFingerprint);
        Assert.Equal("verified", receipt.Status);
    }

    [Fact]
    public void Compare_FingerprintOnlyRange_ReturnsPartiallyVerified()
    {
        var before = Snapshot("same", "rangeFingerprint", ["same"], bounded: true);
        var after = Snapshot("same", "rangeFingerprint", ["same"], bounded: true);

        var receipt = WorkbookSemanticInspector.Compare(before, after);

        Assert.Equal("partiallyVerified", receipt.Status);
        Assert.NotNull(receipt.Limitation);
    }

    [Fact]
    public void Compare_OpaqueOperation_ReturnsNotVerified()
    {
        var before = Snapshot("before", "notVerified", [], bounded: true);
        var after = Snapshot("after", "notVerified", [], bounded: true);

        var receipt = WorkbookSemanticInspector.Compare(before, after);

        Assert.Equal("notVerified", receipt.Status);
        Assert.NotNull(receipt.Limitation);
    }

    [Theory]
    [InlineData("{\"sheetName\":\"Data\",\"rangeAddress\":\"A1:B2\",\"tableName\":\"Sales\",\"chartName\":\"Revenue\",\"pivotTableName\":\"Summary\",\"queryName\":\"Import\",\"connectionName\":\"Warehouse\",\"name\":\"TaxRate\"}")]
    public void ResolveTarget_CollectsKnownStructuralIdentifiers(string argsJson)
    {
        var target = WorkbookSemanticInspector.ResolveTarget(argsJson);

        Assert.Equal("Data", target.SheetName);
        Assert.Equal("A1:B2", target.RangeAddress);
        Assert.Equal("Sales", target.TableName);
        Assert.Equal("Revenue", target.ChartName);
        Assert.Equal("Summary", target.PivotTableName);
        Assert.Equal("Import", target.QueryName);
        Assert.Equal("Warehouse", target.ConnectionName);
        Assert.Equal("TaxRate", target.Name);
    }

    [Theory]
    [InlineData("worksheet", "{\"sheetName\":\"Data\"}", "worksheet:Data")]
    [InlineData("table", "{\"tableName\":\"Sales\"}", "table:Sales")]
    [InlineData("chart", "{\"chartName\":\"Revenue\"}", "chart:Revenue")]
    [InlineData("pivotTable", "{\"pivotTableName\":\"Summary\"}", "pivotTable:Summary")]
    [InlineData("externalObject", "{\"queryName\":\"Import\"}", "powerQuery:Import")]
    [InlineData("workbook", "{\"name\":\"TaxRate\"}", "name:TaxRate")]
    public void ResolveScope_ReportsResolvableObjectIdentifier(string resolver, string argsJson, string expectedObject)
    {
        var scope = WorkbookSemanticInspector.ResolveScope(resolver, WorkbookSemanticInspector.ResolveTarget(argsJson));

        Assert.Contains(expectedObject, scope.Objects);
    }

    [Fact]
    public void Compare_StructuralSemanticEvidence_RemainsPartiallyVerified()
    {
        var before = Snapshot("before", "collectionSemantic", [], bounded: true);
        var after = Snapshot("after", "collectionSemantic", [], bounded: true);

        var receipt = WorkbookSemanticInspector.Compare(before, after);

        Assert.Equal("partiallyVerified", receipt.Status);
        Assert.Equal(0, receipt.ChangedCells);
    }

    [Fact]
    public void Compare_RangeSemanticWithoutCapturedCells_RemainsPartiallyVerified()
    {
        var before = Snapshot("before", "rangeSemantic", [], bounded: true);
        var after = Snapshot("after", "rangeSemantic", [], bounded: true);

        var receipt = WorkbookSemanticInspector.Compare(before, after);

        Assert.Equal("partiallyVerified", receipt.Status);
    }

    [Theory]
    [InlineData(unchecked((int)0x800706BA))]
    [InlineData(unchecked((int)0x800706BE))]
    [InlineData(unchecked((int)0x80010108))]
#pragma warning disable CA2201 // COMException is the production surface being classified.
    public void IsInspectableFailure_ExcelDisconnectsAreNotSwallowed(int hresult)
    {
        var exception = new COMException("Excel disconnected", hresult);

        Assert.False(WorkbookSemanticInspector.IsInspectableFailure(exception));
    }
#pragma warning restore CA2201

    [Fact]
    public void ResolveTarget_EmptySheetNamePreservesWorkbookNamedRangeAddress()
    {
        var target = WorkbookSemanticInspector.ResolveTarget("{\"sheetName\":\"\",\"rangeAddress\":\"InputCell\"}");

        Assert.Equal(string.Empty, target.SheetName);
        Assert.Equal("InputCell", target.RangeAddress);
    }

    [Theory]
    [InlineData("verified", "verified", null)]
    [InlineData("partiallyVerified", "partiallyVerified", null)]
    [InlineData("notVerified", "notVerified", null)]
    [InlineData("failed", "verificationFailed", "VerificationFailed")]
    public void ClassifyVerificationTransition_UsesHonestTerminalState(
        string status,
        string expectedState,
        string? expectedCategory)
    {
        var (state, category) = WorkbookSafetyCoordinator.ClassifyVerificationTransition(status);

        Assert.Equal(expectedState, state);
        Assert.Equal(expectedCategory, category);
    }

    [Fact]
    public void JournalPersistenceError_ReportsMutationExecutedAndAmbiguous()
    {
        var response = WorkbookSafetyCoordinator.JournalPersistenceError(
            "operation-123",
            "durable evidence unavailable");

        Assert.False(response.Success);
        Assert.Equal("JournalPersistenceFailed", response.ErrorCategory);
        Assert.Contains("durable evidence", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("operation-123", response.Result, StringComparison.Ordinal);
        Assert.Contains("completed-but-durable-evidence-unavailable", response.Result, StringComparison.Ordinal);
    }

    private static SemanticSnapshot Snapshot(
        string fingerprint,
        string verificationLevel,
        IReadOnlyList<string> cells,
        bool bounded,
        SafetyScope? scope = null) => new(
            fingerprint,
            fingerprint,
            scope ?? SafetyScope.Workbook,
            verificationLevel,
            bounded,
            cells,
            cells.Count);
}
