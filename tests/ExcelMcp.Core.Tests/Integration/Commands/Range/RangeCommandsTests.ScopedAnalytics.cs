using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

public partial class RangeCommandsTests
{
    [Fact]
    public void SampleValues_ReturnsBoundaryRowsWithSourceCoordinates()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(batch, sheetName, "B2:C6",
        [
            ["R1B", "R1C"],
            ["R2B", "R2C"],
            ["R3B", "R3C"],
            ["R4B", "R4C"],
            ["R5B", "R5C"]
        ]);

        var result = _commands.SampleValues(batch, sheetName, "B2:C6", 2, 1);

        Assert.True(result.Success);
        Assert.Equal(3, result.Rows.Count);
        Assert.Equal(0, result.Rows[0].RowOffset);
        Assert.Equal(2, result.Rows[0].RowNumber);
        Assert.Equal(4, result.Rows[2].RowOffset);
        Assert.Equal(6, result.Rows[2].RowNumber);
        Assert.Equal("R5C", result.Rows[2].Values[1]);
    }

    [Fact]
    public void SummarizeValues_ReturnsTypedCountsAndNumericStatistics()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(batch, sheetName, "A1:A4", [[1], [3], ["text"], [true]]);
        _commands.SetFormulas(batch, sheetName, "A6", [["=1/0"]]);

        var result = _commands.SummarizeValues(batch, sheetName, "A1:A6");
        var summary = result.Columns[0];

        Assert.True(result.Success);
        Assert.Equal(6L, summary.CellCount);
        Assert.Equal(1L, summary.BlankCount);
        Assert.Equal(2L, summary.NumericCount);
        Assert.Equal(1L, summary.TextCount);
        Assert.Equal(1L, summary.LogicalCount);
        Assert.Equal(1L, summary.ErrorCount);
        Assert.Equal(4d, summary.Sum);
        Assert.Equal(2d, summary.Average);
        Assert.Equal(1d, summary.Minimum);
        Assert.Equal(3d, summary.Maximum);
    }

    [Fact]
    public void GetFormulaErrors_ReturnsSparseBoundedDiagnostics()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetFormulas(batch, sheetName, "B2:B4", [["=1+1"], ["=1/0"], ["=NA()"]]);

        var result = _commands.GetFormulaErrors(batch, sheetName, "B2:B4", 1);
        var error = result.Errors[0];

        Assert.True(result.Success);
        Assert.Equal(2L, result.TotalErrorCount);
        Assert.Equal(1, result.ReturnedErrorCount);
        Assert.True(result.IsTruncated);
        Assert.Equal("B3", error.CellAddress);
        Assert.Equal("=1/0", error.Formula);
        Assert.Contains("#DIV/0!", error.ErrorMessage, StringComparison.Ordinal);

        var allErrors = _commands.GetFormulaErrors(batch, sheetName, "B2:B4");
        Assert.Contains(allErrors.Errors, candidate =>
            candidate.Formula == "=NA()" &&
            candidate.ErrorMessage.Contains("#N/A", StringComparison.Ordinal));
    }

    [Fact]
    public void GetFormulaErrors_WithSingleCell_DoesNotEscapeRequestedRange()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetFormulas(batch, sheetName, "A1", [["=1/0"]]);
        _commands.SetFormulas(batch, sheetName, "B2", [["=1+1"]]);

        var result = _commands.GetFormulaErrors(batch, sheetName, "B2");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(0, result.TotalErrorCount);
        Assert.Empty(result.Errors);
    }

    [Fact]
    public void GetFormulaErrors_ExcludesConstantErrorValues()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetFormulas(batch, sheetName, "C1", [["=1/0"]]);
        _commands.CopyValues(batch, sheetName, "C1", sheetName, "D1");

        var result = _commands.GetFormulaErrors(batch, sheetName, "C1:D1");

        Assert.True(result.Success, result.ErrorMessage);
        var error = Assert.Single(result.Errors);
        Assert.Equal("C1", error.CellAddress);
        Assert.Equal("=1/0", error.Formula);
    }

    [Fact]
    public void SampleValues_WithNamedRangeAndSelectedColumns_DeduplicatesOverlap()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var name = $"Sample_{Guid.NewGuid():N}";
        var namedRanges = new NamedRangeCommands();
        namedRanges.Create(batch, name, $"{sheetName}!$C$3:$E$5");
        try
        {
            _commands.SetValues(batch, sheetName, "C3:E5",
            [
                ["R1C", "R1D", "R1E"],
                ["R2C", "R2D", "R2E"],
                ["R3C", "R3D", "R3E"]
            ]);

            var result = _commands.SampleValues(batch, "", name, 2, 2, "E,C");

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal(3, result.Rows.Count);
            Assert.Equal(["E", "C"], result.SelectedColumns);
            Assert.Equal([0, 1, 2], result.Rows.Select(row => row.RowOffset));
            Assert.Equal(["R3E", "R3C"], result.Rows[2].Values);
            Assert.Equal("$E$5,$C$5", result.Rows[2].RangeAddress);
        }
        finally
        {
            namedRanges.Delete(batch, name);
        }
    }

    [Fact]
    public void SummarizeValues_WithSelectedColumns_PreservesOrderAndNullNumericStatistics()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var name = $"Summary_{Guid.NewGuid():N}";
        var namedRanges = new NamedRangeCommands();
        _commands.SetValues(batch, sheetName, "B2:D4",
        [
            ["b1", 10, true],
            ["b2", 20, false],
            ["b3", 30, true]
        ]);
        namedRanges.Create(batch, name, $"{sheetName}!$B$2:$D$4");

        try
        {
            var result = _commands.SummarizeValues(batch, "", name, "D,B");

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal(["D", "B"], result.SelectedColumns);
            Assert.Equal(["D", "B"], result.Columns.Select(column => column.Column));
            Assert.Equal(3, result.Columns[0].LogicalCount);
            Assert.Null(result.Columns[0].Sum);
            Assert.Equal(3, result.Columns[1].TextCount);
            Assert.Null(result.Columns[1].Average);
        }
        finally
        {
            namedRanges.Delete(batch, name);
        }
    }

    [Fact]
    public void SummarizeValues_WithLargeSourceRange_AggregatesWithoutReturningValues()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(batch, sheetName, "A99999:A100000", [[2], [4]]);

        var result = _commands.SummarizeValues(batch, sheetName, "A1:A100000");
        var summary = Assert.Single(result.Columns);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(100000, summary.CellCount);
        Assert.Equal(99998, summary.BlankCount);
        Assert.Equal(2, summary.NumericCount);
        Assert.Equal(6, summary.Sum);
    }

    [Fact]
    public void SummarizeValues_WithFragmentedTypes_ReturnsCompleteCounts()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var values = Enumerable.Range(0, 17000)
            .Select(index => new List<object?> { index % 2 == 0 ? 1 : "text" })
            .ToList();
        _commands.SetValues(batch, sheetName, "A1:A17000", values);

        var result = _commands.SummarizeValues(batch, sheetName, "A1:A17000");
        var summary = Assert.Single(result.Columns);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(8500, summary.NumericCount);
        Assert.Equal(8500, summary.TextCount);
        Assert.Equal(8500, summary.Sum);
    }

    [Fact]
    public void GetFormulaErrors_WithFragmentedResults_ReturnsCompleteCount()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var formulas = Enumerable.Range(0, 17000)
            .Select(index => new List<string> { index % 2 == 0 ? "=1/0" : "=1+1" })
            .ToList();
        _commands.SetFormulas(batch, sheetName, "A1:A17000", formulas);

        var result = _commands.GetFormulaErrors(batch, sheetName, "A1:A17000", 1);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(8500, result.TotalErrorCount);
        Assert.Single(result.Errors);
        Assert.True(result.IsTruncated);
    }

    [Fact]
    public void GetFormulaErrors_WithMultiAreaNamedRange_ReturnsWorksheetOrderedDiagnostics()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var name = $"FormulaErrors_{Guid.NewGuid():N}";
        var namedRanges = new NamedRangeCommands();
        _commands.SetFormulas(batch, sheetName, "A1", [["=1/0"]]);
        _commands.SetFormulas(batch, sheetName, "C2", [["=NA()"]]);

        batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? firstArea = null;
            Excel.Range? secondArea = null;
            Excel.Range? unionRange = null;
            Excel.Names? names = null;
            Excel.Name? namedRange = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                firstArea = sheet.Range["C1:C2"];
                secondArea = sheet.Range["A1:A2"];
                unionRange = ctx.App.Union(firstArea, secondArea);
                names = ctx.Book.Names;
                namedRange = names.Add(name, unionRange);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref namedRange);
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref unionRange);
                ComUtilities.Release(ref secondArea);
                ComUtilities.Release(ref firstArea);
                ComUtilities.Release(ref sheet);
            }
        });

        try
        {
            var result = _commands.GetFormulaErrors(batch, "", name);

            Assert.True(result.Success, result.ErrorMessage);
            Assert.Equal(2, result.TotalErrorCount);
            Assert.Equal(["A1", "C2"], result.Errors.Select(error => error.CellAddress));
            Assert.Equal(["=1/0", "=NA()"], result.Errors.Select(error => error.Formula));
        }
        finally
        {
            namedRanges.Delete(batch, name);
        }
    }

    [Theory]
    [InlineData(-1, 1, "firstRowCount")]
    [InlineData(1, -1, "lastRowCount")]
    [InlineData(101, 1, "firstRowCount")]
    [InlineData(1, 101, "lastRowCount")]
    [InlineData(0, 0, "firstRowCount")]
    public void SampleValues_WithInvalidCounts_ThrowsBeforeOpeningExcel(
        int firstRowCount,
        int lastRowCount,
        string parameterName)
    {
        var exception = Assert.ThrowsAny<ArgumentException>(() =>
            _commands.SampleValues(
                null!,
                "Sheet1",
                "A1",
                firstRowCount,
                lastRowCount));

        Assert.Equal(parameterName, exception.ParamName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1001)]
    public void GetFormulaErrors_WithInvalidLimit_ThrowsBeforeOpeningExcel(int maxErrors)
    {
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            _commands.GetFormulaErrors(null!, "Sheet1", "A1", maxErrors));

        Assert.Equal("maxErrors", exception.ParamName);
    }

    [Fact]
    public void SampleValues_WhenResponseWouldExceedCellLimit_Throws()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.SampleValues(batch, sheetName, "A1:Z100", 50, 50));

        Assert.Contains("1000 cells", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SamplingAndSummaries_WithMultiAreaRange_RejectClearly()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        var name = $"ScopedModes_{Guid.NewGuid():N}";
        var namedRanges = new NamedRangeCommands();

        batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? firstArea = null;
            Excel.Range? secondArea = null;
            Excel.Range? unionRange = null;
            Excel.Names? names = null;
            Excel.Name? namedRange = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                firstArea = sheet.Range["A1:A2"];
                secondArea = sheet.Range["C1:C2"];
                unionRange = ctx.App.Union(firstArea, secondArea);
                names = ctx.Book.Names;
                namedRange = names.Add(name, unionRange);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref namedRange);
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref unionRange);
                ComUtilities.Release(ref secondArea);
                ComUtilities.Release(ref firstArea);
                ComUtilities.Release(ref sheet);
            }
        });

        try
        {
            var sampleException = Assert.Throws<ArgumentException>(() =>
                _commands.SampleValues(batch, "", name));
            var summaryException = Assert.Throws<ArgumentException>(() =>
                _commands.SummarizeValues(batch, "", name));

            Assert.Contains("multi-area", sampleException.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("multi-area", summaryException.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            namedRanges.Delete(batch, name);
        }
    }
}
