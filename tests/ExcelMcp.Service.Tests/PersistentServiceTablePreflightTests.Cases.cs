using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceTablePreflightTests
{
    [Fact]
    public void Preflight_SingleCellInsideData_ReportsExpandedEffectiveRange()
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "F1:H3",
        [
            ["Name", "Region", "Amount"],
            ["Widget", "North", 10],
            ["Gadget", "South", 20]
        ]);
        var before = GetSourceState("Sales", "F1:H3");

        var result = _tableCommands.Preflight(batch, "Sales", "ExpandedTable", "G2");

        RequireSuccess(result);
        Assert.True(result.SafeToCreate);
        Assert.Equal("G2", result.RequestedRange);
        Assert.Equal("$F$1:$H$3", result.EffectiveRange);
        Assert.Empty(result.Findings);
        AssertPreflightPreserved(before, "F1:H3");
    }

    [Fact]
    public void Preflight_MergedCells_ReturnsBlockerAndCreateDoesNotChangeWorkbook()
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "F1:H3",
        [
            ["Name", "Region", "Amount"],
            ["Widget", "North", 10],
            ["Gadget", "South", 20]
        ]);
        RequireSuccess(_rangeCommands.MergeCells(batch, "Sales", "G2:H2"));
        var before = GetSourceState("Sales", "F1:H3");

        var result = _tableCommands.Preflight(batch, "Sales", "MergedTable", "F1:H3");

        RequireSuccess(result);
        AssertPreflightPreserved(before, "F1:H3");
        Assert.False(result.SafeToCreate);
        var finding = Assert.Single(result.Findings, item => item.Kind == TablePreflightFindingKind.MergedCells);
        Assert.Equal(TablePreflightSeverity.Blocker, finding.Severity);
        Assert.False(finding.IsHeuristic);
        Assert.Contains("$G$2:$H$2", finding.Addresses);
        Assert.False(string.IsNullOrWhiteSpace(finding.Remediation));

        var exception = Assert.Throws<InvalidOperationException>(
            () => _tableCommands.Create(batch, "Sales", "MergedTable", "F1:H3"));
        Assert.Contains("merged", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertPreflightPreserved(before, "F1:H3");
        var merge = RequireSuccess(_rangeCommands.GetMergeInfo(batch, "Sales", "G2:H2"));
        Assert.True(merge.IsMerged);
    }

    [Fact]
    public void Preflight_BlankAndDuplicateHeaders_ReturnsAddressedBlockers()
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "F1:H2",
        [
            [null, "Name", " name "],
            [1, "Widget", "Duplicate"]
        ]);
        var before = GetSourceState("Sales", "F1:H2");

        var result = _tableCommands.Preflight(batch, "Sales", "HeaderTable", "F1:H2");

        RequireSuccess(result);
        AssertPreflightPreserved(before, "F1:H2");
        Assert.False(result.SafeToCreate);
        var blank = Assert.Single(result.Findings, item => item.Kind == TablePreflightFindingKind.BlankHeaders);
        Assert.Equal(TablePreflightSeverity.Blocker, blank.Severity);
        Assert.Equal(["$F$1"], blank.Addresses);

        var duplicate = Assert.Single(result.Findings, item => item.Kind == TablePreflightFindingKind.DuplicateHeaders);
        Assert.Equal(TablePreflightSeverity.Blocker, duplicate.Severity);
        Assert.Equal(["$G$1", "$H$1"], duplicate.Addresses);

        Assert.Throws<InvalidOperationException>(
            () => _tableCommands.Create(batch, "Sales", "HeaderTable", "F1:H2"));
        AssertPreflightPreserved(before, "F1:H2");
    }

    [Fact]
    public void Preflight_ExcludedContiguousColumn_ReturnsNonBlockingWarning()
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "F1:H3",
        [
            ["Name", "Region", "Amount"],
            ["Widget", "North", 10],
            ["Gadget", "South", 20]
        ]);
        var before = GetSourceState("Sales", "F1:H3");

        var result = _tableCommands.Preflight(batch, "Sales", "NarrowTable", "F1:G3");

        RequireSuccess(result);
        AssertPreflightPreserved(before, "F1:H3");
        Assert.True(result.SafeToCreate);
        var finding = Assert.Single(
            result.Findings,
            item => item.Kind == TablePreflightFindingKind.ExcludedContiguousColumns);
        Assert.Equal(TablePreflightSeverity.Warning, finding.Severity);
        Assert.True(finding.IsHeuristic);
        Assert.Equal(["$H$1:$H$3"], finding.Addresses);

        RequireSuccess(_tableCommands.Create(batch, "Sales", "NarrowTable", "F1:G3"));
        Assert.Equal("$F$1:$G$3", RequireSuccess(_tableCommands.Read(batch, "NarrowTable")).Table!.Range);
        var data = RequireSuccess(_tableCommands.GetData(batch, "NarrowTable"));
        Assert.Equal(["Name", "Region"], data.Headers);
        Assert.Equal(["Widget", "North"], data.Data[0]);
        Assert.Equal(["Gadget", "South"], data.Data[1]);
        Assert.Equal(2, data.Data.Count);
        Assert.Equal(before.Formulas.Cast<object>(), GetSourceState("Sales", "F1:H3").Formulas.Cast<object>());
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
    }

    [Fact]
    public void Preflight_SortSensitiveFormula_ReturnsNonBlockingWarning()
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "F1:G3",
        [
            ["Amount", "Calculated"],
            [10, null],
            [20, 40]
        ]);
        RequireSuccess(_rangeCommands.SetFormulas(
            batch,
            "Sales",
            "G2:G3",
            [
                ["=$F$2*2"],
                ["=I3*2"]
            ],
            overwritePolicy: OverwritePolicy.Allow));
        var before = GetSourceState("Sales", "F1:I3");

        var result = _tableCommands.Preflight(batch, "Sales", "FormulaTable", "F1:G3");

        RequireSuccess(result);
        AssertPreflightPreserved(before, "F1:I3");
        Assert.True(result.SafeToCreate);
        var finding = Assert.Single(
            result.Findings,
            item => item.Kind == TablePreflightFindingKind.SortSensitiveFormula);
        Assert.Equal(TablePreflightSeverity.Warning, finding.Severity);
        Assert.True(finding.IsHeuristic);
        Assert.Equal(["$G$2", "$G$3"], finding.Addresses);

        RequireSuccess(_tableCommands.Create(batch, "Sales", "FormulaTable", "F1:G3"));
        Assert.Equal("$F$1:$G$3", RequireSuccess(_tableCommands.Read(batch, "FormulaTable")).Table!.Range);
        var formulas = RequireSuccess(_rangeCommands.GetFormulas(batch, "Sales", "G2:G3"));
        Assert.Equal("=$F$2*2", formulas.Formulas[0][0]);
        Assert.Equal("=I3*2", formulas.Formulas[1][0]);
        var values = RequireSuccess(_rangeCommands.GetValues(batch, "Sales", "G2:G3"));
        Assert.Equal([20d, 0d], values.Values.Select(row => Convert.ToDouble(row[0], System.Globalization.CultureInfo.InvariantCulture)));
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
    }

    [Fact]
    public void Preflight_ExistingTableName_ReturnsBlocker()
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "F1:G2",
        [
            ["Name", "Amount"],
            ["Widget", 10]
        ]);
        var before = GetSourceState("Sales", "F1:G2");

        var result = _tableCommands.Preflight(batch, "Sales", "SalesTable", "F1:G2");

        RequireSuccess(result);
        AssertPreflightPreserved(before, "F1:G2");
        Assert.False(result.SafeToCreate);
        var finding = Assert.Single(
            result.Findings,
            item => item.Kind == TablePreflightFindingKind.TableNameExists);
        Assert.Equal(TablePreflightSeverity.Blocker, finding.Severity);
        Assert.Empty(finding.Addresses);
        Assert.Contains("already exists", finding.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Throws<InvalidOperationException>(() => _tableCommands.Create(batch, "Sales", "SalesTable", "F1:G2"));
        AssertPreflightPreserved(before, "F1:G2");
    }

    [Fact]
    public void Preflight_WithoutHeaders_DoesNotReportBlankHeaderBlocker()
    {
        var batch = _fixture.BatchToken;
        SetValues(batch, "F1:G2",
        [
            [10, null],
            [20, 40]
        ]);
        RequireSuccess(_rangeCommands.SetFormulas(batch, "Sales", "G1", [["=$F$1*2"]]));
        var before = GetSourceState("Sales", "F1:G2");

        var result = _tableCommands.Preflight(
            batch,
            "Sales",
            "HeaderlessTable",
            "F1:G2",
            hasHeaders: false);

        RequireSuccess(result);
        AssertPreflightPreserved(before, "F1:G2");
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
        var batch = _fixture.BatchToken;
        var before = GetSourceState("Sales", "A1:H5");

        var result = _tableCommands.Preflight(
            batch,
            "Sales",
            "LargeTable",
            "A1:CV1001",
            hasHeaders: false);

        RequireSuccess(result);
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
        AssertPreflightPreserved(before, "A1:H5");
    }

    private void SetValues(IExcelBatch batch, string address, List<List<object?>> values)
    {
        RequireSuccess(_rangeCommands.SetValues(batch, "Sales", address, values));
    }

    private void AssertPreflightPreserved(
        (object[,] Formulas, string[] Formats, int WorkbookCount) before, string address)
    {
        var after = GetSourceState("Sales", address);
        Assert.Equal(before.Formulas.Cast<object>(), after.Formulas.Cast<object>());
        Assert.Equal(before.Formats, after.Formats);
        Assert.Equal(before.WorkbookCount, after.WorkbookCount);
        AssertSalesTableInfo(Assert.Single(RequireSuccess(_tableCommands.List(_fixture.BatchToken)).Tables));
        AssertSalesData(RequireSuccess(_tableCommands.GetData(_fixture.BatchToken, "SalesTable")));
    }
}
