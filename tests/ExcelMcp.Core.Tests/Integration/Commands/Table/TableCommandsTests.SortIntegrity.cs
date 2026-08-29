using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

public partial class TableCommandsTests
{
    [Fact]
    public void Sort_DefaultBehavior_SortsWithoutIntegrityValidation()
    {
        var testFile = CopyTableFixture();
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _tableCommands.Sort(
            batch,
            "SalesTable",
            "Amount",
            ascending: true,
            validateIntegrity: false);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(result.ValidationPerformed);
        Assert.Null(result.IntegrityPreserved);
        Assert.True(result.SortAttempted);
        Assert.True(result.SortCommitted);
        Assert.False(result.RollbackAttempted);
        Assert.Null(result.RollbackSucceeded);

        var data = _tableCommands.GetData(batch, "SalesTable", visibleOnly: false).Data;
        Assert.Equal(["North", "East", "South", "West"], data.Select(row => row[0]?.ToString()));
        Assert.Equal(["100", "150", "250", "300"], data.Select(row => row[2]?.ToString()));
    }

    [Fact]
    public void Sort_WithIntegrityValidation_ProvesCompleteRowsWerePermuted()
    {
        var testFile = CopyTableFixture();
        using var batch = ExcelSession.BeginBatch(testFile);
        List<string> beforeRows = GetCompleteRowSignatures(
            _tableCommands.GetData(batch, "SalesTable", visibleOnly: false).Data);

        var result = _tableCommands.Sort(
            batch,
            "SalesTable",
            "Amount",
            validateIntegrity: true);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.ValidationPerformed);
        Assert.True(result.IntegrityPreserved);
        Assert.True(result.SortCommitted);
        Assert.True(result.Checks.RangePreserved);
        Assert.True(result.Checks.ShapePreserved);
        Assert.True(result.Checks.HeadersPreserved);
        Assert.True(result.Checks.TotalsRowPreserved);
        Assert.True(result.Checks.RowSetPreserved);
        Assert.False(result.RollbackAttempted);
        List<string> afterRows = GetCompleteRowSignatures(
            _tableCommands.GetData(batch, "SalesTable", visibleOnly: false).Data);
        Assert.Equal(beforeRows, afterRows);
    }

    [Fact]
    public void SortMulti_WithIntegrityValidation_UsesSameValidationPipeline()
    {
        var testFile = CopyTableFixture();
        using var batch = ExcelSession.BeginBatch(testFile);

        var result = _tableCommands.SortMulti(
            batch,
            "SalesTable",
            [
                new TableSortColumn { ColumnName = "Product", Ascending = true },
                new TableSortColumn { ColumnName = "Amount", Ascending = false }
            ],
            validateIntegrity: true);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(result.ValidationPerformed);
        Assert.True(result.IntegrityPreserved);
        Assert.True(result.SortCommitted);
        Assert.True(result.Checks.RowSetPreserved);

        var data = _tableCommands.GetData(batch, "SalesTable", visibleOnly: false).Data;
        Assert.Equal(["Gadget", "Gadget", "Widget", "Widget"], data.Select(row => row[1]?.ToString()));
        Assert.Equal(["300", "250", "150", "100"], data.Select(row => row[2]?.ToString()));
    }

    [Fact]
    public void Sort_WithAdjacentDataAndExternalFormulas_ReturnsHeuristicWarnings()
    {
        var testFile = CopyTableFixture();
        using var batch = ExcelSession.BeginBatch(testFile);
        _tableCommands.Delete(batch, "SalesTable");
        _rangeCommands.SetValues(
            batch,
            "Sales",
            "E1",
            [["Related"]]);
        _rangeCommands.SetFormulas(
            batch,
            "Sales",
            "E2:E5",
            [["=C2*2"], ["=C3*2"], ["=C4*2"], ["=C5*2"]]);
        _tableCommands.Create(batch, "Sales", "SalesTable", "A1:D5");

        var result = _tableCommands.Sort(
            batch,
            "SalesTable",
            "Amount",
            validateIntegrity: true);

        Assert.True(result.Success, result.ErrorMessage);
        var adjacent = Assert.Single(
            result.Findings,
            finding => finding.Kind == TablePreflightFindingKind.ExcludedContiguousColumns);
        Assert.Equal(TablePreflightSeverity.Warning, adjacent.Severity);
        Assert.True(adjacent.IsHeuristic);

        var external = Assert.Single(
            result.Findings,
            finding => finding.Kind == TablePreflightFindingKind.RowAssociatedFormulaOutsideTable);
        Assert.Equal(TablePreflightSeverity.Warning, external.Severity);
        Assert.True(external.IsHeuristic);
        Assert.Equal(["$E$2", "$E$3", "$E$4", "$E$5"], external.Addresses);
    }

    [Fact]
    public void Sort_WithDuplicateCompositeKeys_ReturnsBlockerWithoutMutation()
    {
        var testFile = CopyTableFixture();
        using var batch = ExcelSession.BeginBatch(testFile);
        var before = _tableCommands.GetData(batch, "SalesTable", visibleOnly: false).Data;

        var result = _tableCommands.Sort(
            batch,
            "SalesTable",
            "Amount",
            validateIntegrity: true,
            keyColumns: ["Product"]);

        Assert.False(result.Success);
        Assert.False(result.SortAttempted);
        Assert.False(result.SortCommitted);
        Assert.Null(result.IntegrityPreserved);
        var blocker = Assert.Single(
            result.Findings,
            finding => finding.Kind == TablePreflightFindingKind.DuplicateRowKey);
        Assert.Equal(TablePreflightSeverity.Blocker, blocker.Severity);
        Assert.False(blocker.IsHeuristic);
        Assert.Equal(before, _tableCommands.GetData(batch, "SalesTable", visibleOnly: false).Data);
    }

    [Fact]
    public void Sort_WithCalculatedColumn_VerifiesFormulaConsistency()
    {
        var testFile = CopyTableFixture();
        using var batch = ExcelSession.BeginBatch(testFile);
        _tableCommands.AddColumn(batch, "SalesTable", "DoubleAmount");
        _rangeCommands.SetFormulas(
            batch,
            "Sales",
            "E2:E5",
            [
                ["=[@Amount]*2"],
                ["=[@Amount]*2"],
                ["=[@Amount]*2"],
                ["=[@Amount]*2"]
            ]);

        var result = _tableCommands.Sort(
            batch,
            "SalesTable",
            "Amount",
            validateIntegrity: true);

        Assert.True(result.Success, result.ErrorMessage);
        var calculated = Assert.Single(
            result.Checks.CalculatedColumns,
            check => check.ColumnName == "DoubleAmount");
        Assert.True(calculated.ConsistentBefore);
        Assert.True(calculated.ConsistentAfter);
        Assert.True(calculated.Passed);
    }

    [Fact]
    public void Sort_WhenControlTotalChanges_RestoresSnapshotAndReturnsTypedFailure()
    {
        var testFile = CopyTableFixture();
        using var batch = ExcelSession.BeginBatch(testFile);
        _tableCommands.AddColumn(batch, "SalesTable", "Weighted");
        _tableCommands.AddColumn(batch, "SalesTable", "FormulaLikeText");
        _rangeCommands.SetFormulas(
            batch,
            "Sales",
            "E2:E5",
            [
                ["=[@Amount]*ROW()"],
                ["=[@Amount]*ROW()"],
                ["=[@Amount]*ROW()"],
                ["=[@Amount]*ROW()"]
            ]);
        _rangeCommands.SetValues(
            batch,
            "Sales",
            "F2:F5",
            [["'=1+1"], ["'=2+2"], ["'=3+3"], ["'=4+4"]]);
        var before = _tableCommands.GetData(batch, "SalesTable", visibleOnly: false).Data;

        var result = _tableCommands.Sort(
            batch,
            "SalesTable",
            "Amount",
            validateIntegrity: true,
            keyColumns: ["Region"],
            controlTotals:
            [
                new TableSortControlTotal { ColumnName = "Weighted" }
            ]);

        Assert.False(result.Success);
        Assert.True(result.SortAttempted);
        Assert.False(result.SortCommitted);
        Assert.False(result.IntegrityPreserved);
        Assert.True(result.RollbackAttempted);
        Assert.True(result.RollbackSucceeded);
        var total = Assert.Single(result.Checks.ControlTotals);
        Assert.False(total.Passed);
        Assert.NotEqual(total.Before, total.After);
        Assert.Contains(
            result.Findings,
            finding => finding.Kind == TablePreflightFindingKind.ControlTotalMismatch);
        Assert.Equal(before, _tableCommands.GetData(batch, "SalesTable", visibleOnly: false).Data);
        Assert.False(batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets["Sales"];
                cell = sheet.Range["F2"];
                return Convert.ToBoolean(cell.HasFormula);
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        }));
    }

    private string CopyTableFixture()
    {
        string path = Path.Join(_fixture.TempDir, $"SortIntegrity_{Guid.NewGuid():N}.xlsx");
        System.IO.File.Copy(_tableFile, path);
        return path;
    }

    private static List<string> GetCompleteRowSignatures(List<List<object?>> rows) =>
        rows.Select(row => string.Join(
                "\u001f",
                row.Select(value => Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture))))
            .OrderBy(signature => signature, StringComparer.Ordinal)
            .ToList();
}
