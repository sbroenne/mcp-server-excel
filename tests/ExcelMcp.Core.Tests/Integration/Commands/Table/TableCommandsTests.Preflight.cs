using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

public partial class TableCommandsTests
{
    [Fact]
    public void Preflight_SingleCellInsideData_ReportsExpandedEffectiveRange()
    {
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        SetValues(batch, "F1:H3",
        [
            ["Name", "Region", "Amount"],
            ["Widget", "North", 10],
            ["Gadget", "South", 20]
        ]);

        var result = _tableCommands.Preflight(batch, "Sales", "ExpandedTable", "G2");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.SafeToCreate);
        Assert.Equal("G2", result.RequestedRange);
        Assert.Equal("$F$1:$H$3", result.EffectiveRange);
        Assert.Empty(result.Findings);
    }

    [Fact]
    public void Preflight_MergedCells_ReturnsBlockerAndCreateDoesNotChangeWorkbook()
    {
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        SetValues(batch, "F1:H3",
        [
            ["Name", "Region", "Amount"],
            ["Widget", "North", 10],
            ["Gadget", "South", 20]
        ]);
        _rangeCommands.MergeCells(batch, "Sales", "G2:H2");

        var result = _tableCommands.Preflight(batch, "Sales", "MergedTable", "F1:H3");

        Assert.False(result.SafeToCreate);
        var finding = Assert.Single(result.Findings, item => item.Kind == TablePreflightFindingKind.MergedCells);
        Assert.Equal(TablePreflightSeverity.Blocker, finding.Severity);
        Assert.False(finding.IsHeuristic);
        Assert.Contains("$G$2:$H$2", finding.Addresses);
        Assert.False(string.IsNullOrWhiteSpace(finding.Remediation));

        var exception = Assert.Throws<InvalidOperationException>(
            () => _tableCommands.Create(batch, "Sales", "MergedTable", "F1:H3"));
        Assert.Contains("merged", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(_tableCommands.List(batch).Tables, table => table.Name == "MergedTable");
        Assert.True(_rangeCommands.GetMergeInfo(batch, "Sales", "G2:H2").IsMerged);
    }

    [Fact]
    public void Preflight_BlankAndDuplicateHeaders_ReturnsAddressedBlockers()
    {
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        SetValues(batch, "F1:H2",
        [
            [null, "Name", " name "],
            [1, "Widget", "Duplicate"]
        ]);

        var result = _tableCommands.Preflight(batch, "Sales", "HeaderTable", "F1:H2");

        Assert.False(result.SafeToCreate);
        var blank = Assert.Single(result.Findings, item => item.Kind == TablePreflightFindingKind.BlankHeaders);
        Assert.Equal(TablePreflightSeverity.Blocker, blank.Severity);
        Assert.Equal(["$F$1"], blank.Addresses);

        var duplicate = Assert.Single(result.Findings, item => item.Kind == TablePreflightFindingKind.DuplicateHeaders);
        Assert.Equal(TablePreflightSeverity.Blocker, duplicate.Severity);
        Assert.Equal(["$G$1", "$H$1"], duplicate.Addresses);

        Assert.Throws<InvalidOperationException>(
            () => _tableCommands.Create(batch, "Sales", "HeaderTable", "F1:H2"));
        Assert.DoesNotContain(_tableCommands.List(batch).Tables, table => table.Name == "HeaderTable");
    }

    [Fact]
    public void Preflight_ExcludedContiguousColumn_ReturnsNonBlockingWarning()
    {
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        SetValues(batch, "F1:H3",
        [
            ["Name", "Region", "Amount"],
            ["Widget", "North", 10],
            ["Gadget", "South", 20]
        ]);

        var result = _tableCommands.Preflight(batch, "Sales", "NarrowTable", "F1:G3");

        Assert.True(result.SafeToCreate);
        var finding = Assert.Single(
            result.Findings,
            item => item.Kind == TablePreflightFindingKind.ExcludedContiguousColumns);
        Assert.Equal(TablePreflightSeverity.Warning, finding.Severity);
        Assert.True(finding.IsHeuristic);
        Assert.Equal(["$H$1:$H$3"], finding.Addresses);

        _tableCommands.Create(batch, "Sales", "NarrowTable", "F1:G3");
        Assert.Equal("$F$1:$G$3", _tableCommands.Read(batch, "NarrowTable").Table!.Range);
    }

    [Fact]
    public void Preflight_SortSensitiveFormula_ReturnsNonBlockingWarning()
    {
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        SetValues(batch, "F1:G3",
        [
            ["Amount", "Calculated"],
            [10, null],
            [20, 40]
        ]);
        _rangeCommands.SetFormulas(
            batch,
            "Sales",
            "G2:G3",
            [
                ["=$F$2*2"],
                ["=I3*2"]
            ]);

        var result = _tableCommands.Preflight(batch, "Sales", "FormulaTable", "F1:G3");

        Assert.True(result.SafeToCreate);
        var finding = Assert.Single(
            result.Findings,
            item => item.Kind == TablePreflightFindingKind.SortSensitiveFormula);
        Assert.Equal(TablePreflightSeverity.Warning, finding.Severity);
        Assert.True(finding.IsHeuristic);
        Assert.Equal(["$G$2", "$G$3"], finding.Addresses);

        _tableCommands.Create(batch, "Sales", "FormulaTable", "F1:G3");
        Assert.Contains(_tableCommands.List(batch).Tables, table => table.Name == "FormulaTable");
    }

    [Fact]
    public void Preflight_ExistingTableName_ReturnsBlocker()
    {
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        SetValues(batch, "F1:G2",
        [
            ["Name", "Amount"],
            ["Widget", 10]
        ]);

        var result = _tableCommands.Preflight(batch, "Sales", "SalesTable", "F1:G2");

        Assert.False(result.SafeToCreate);
        var finding = Assert.Single(
            result.Findings,
            item => item.Kind == TablePreflightFindingKind.TableNameExists);
        Assert.Equal(TablePreflightSeverity.Blocker, finding.Severity);
        Assert.Empty(finding.Addresses);
        Assert.Contains("already exists", finding.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Preflight_WithoutHeaders_DoesNotReportBlankHeaderBlocker()
    {
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        SetValues(batch, "F1:G2",
        [
            [10, null],
            [20, 40]
        ]);
        _rangeCommands.SetFormulas(batch, "Sales", "G1", [["=$F$1*2"]]);

        var result = _tableCommands.Preflight(
            batch,
            "Sales",
            "HeaderlessTable",
            "F1:G2",
            hasHeaders: false);

        Assert.True(result.SafeToCreate);
        Assert.DoesNotContain(
            result.Findings,
            item => item.Kind is TablePreflightFindingKind.BlankHeaders
                or TablePreflightFindingKind.DuplicateHeaders);
        var formulaFinding = Assert.Single(
            result.Findings,
            item => item.Kind == TablePreflightFindingKind.SortSensitiveFormula);
        Assert.Equal(["$G$1"], formulaFinding.Addresses);
    }

    [Fact]
    public void Preflight_OversizedRange_ReturnsExplicitFormulaScanSkippedWarning()
    {
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        var result = _tableCommands.Preflight(
            batch,
            "Sales",
            "LargeTable",
            "A1:CV1001",
            hasHeaders: false);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.SafeToCreate);
        var finding = Assert.Single(
            result.Findings,
            item => item.Kind == TablePreflightFindingKind.FormulaScanSkipped);
        Assert.Equal(TablePreflightSeverity.Warning, finding.Severity);
        Assert.True(finding.IsHeuristic);
        Assert.Empty(finding.Addresses);
        Assert.Contains("100,100", finding.Message, StringComparison.Ordinal);
        Assert.Contains("100,000", finding.Message, StringComparison.Ordinal);
        Assert.Contains("smaller range", finding.Remediation, StringComparison.OrdinalIgnoreCase);
    }

    private void SetValues(IExcelBatch batch, string address, List<List<object?>> values)
    {
        _rangeCommands.SetValues(batch, "Sales", address, values);
    }
}
