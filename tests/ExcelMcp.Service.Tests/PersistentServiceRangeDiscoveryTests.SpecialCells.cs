using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeDiscoveryTests
{
    [Theory]
    [InlineData("formulas", 2, "$A$2", "$A$4")]
    [InlineData("constants", 2, "$A$1", "$A$3")]
    [InlineData("blanks", 2, "$A$5:$A$6")]
    [InlineData("errors", 1, "$A$4")]
    [InlineData("visible", 6, "$A$1:$A$6")]
    public void SpecialCells_ReturnsEveryMatchingAreaInRequestedScope(
        string cellKind, long expectedCount, params string[] expectedAreas)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[42]]).Success);
        Assert.True(_commands.SetFormulas(batch, sheetName, "A2", [["=\"\""]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "A3", [[" "]]).Success);
        Assert.True(_commands.SetFormulas(batch, sheetName, "A4", [["=1/0"]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "C1", [["Outside scope"]]).Success);

        using var result = Discover(sheetName, "A1:A6", cellKind);

        Assert.Equal(expectedCount, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(expectedAreas, ReadAreas(result));
        var values = _commands.GetValues(batch, sheetName, "C1");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal("Outside scope", values.Values[0][0]);
    }

    [Theory]
    [InlineData("A1", "constants", 1)]
    [InlineData("A1", "formulas", 0)]
    [InlineData("A1", "blanks", 0)]
    [InlineData("A1", "errors", 0)]
    [InlineData("A1", "visible", 1)]
    [InlineData("D20", "blanks", 1)]
    [InlineData("D20", "constants", 0)]
    public void SpecialCells_SingleCellDoesNotExpandToUsedRange(
        string rangeAddress, string cellKind, long expectedCount)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[7]]).Success);
        Assert.True(_commands.SetFormulas(batch, sheetName, "B2", [["=1/0"]]).Success);

        using var result = Discover(sheetName, rangeAddress, cellKind);

        Assert.Equal(expectedCount, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(expectedCount == 0 ? [] : [rangeAddress == "A1" ? "$A$1" : "$D$20"],
            ReadAreas(result));
    }

    [Theory]
    [InlineData("formulas")]
    [InlineData("errors")]
    [InlineData("blanks")]
    public void SpecialCells_NoMatchesIsSuccessfulEmptyResult(string cellKind)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A3", [[1], [2], [3]]).Success);

        using var result = Discover(sheetName, "A1:A3", cellKind);

        Assert.Equal(0, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Empty(ReadAreas(result));
    }

    [Fact]
    public void SpecialCells_DoesNotSilentlyLimitAreas()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        List<List<object?>> values = Enumerable.Range(1, 64)
            .Select(row => new List<object?> { row % 2 == 1 ? row : null })
            .ToList();
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A64", values).Success);

        using var result = Discover(sheetName, "A1:A64", "constants");

        Assert.Equal(32, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(Enumerable.Range(0, 32).Select(index => $"$A${index * 2 + 1}"),
            ReadAreas(result));
    }

    [Fact]
    public void SpecialCells_VisibleExcludesHiddenRowsAndColumns()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:C3",
            [[1, 2, 3], [4, 5, 6], [7, 8, 9]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? rows = null;
            Excel.Range? columns = null;
            Excel.Range? row = null;
            Excel.Range? column = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                Assert.False(sheet.ProtectContents);
                rows = sheet.Rows;
                columns = sheet.Columns;
                row = rows[2];
                column = columns[2];
                Assert.Equal("$2:$2", row.Address);
                Assert.Equal("$B:$B", column.Address);
                row.Hidden = true;
                column.Hidden = true;
            }
            finally
            {
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref row);
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref sheet);
            }
        });

        using var result = Discover(sheetName, "A1:C3", "visible");

        Assert.Equal(4, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(["$A$1", "$C$1", "$A$3", "$C$3"], ReadAreas(result));
    }

    [Fact]
    public async Task SpecialCells_InvalidKindIsNotAnEmptySuccess()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var response = await _fixture.SendForFailureAsync("range.get-special-cells",
            new { sheetName, rangeAddress = "A1:A3", cellKind = "unknown-kind" });

        Assert.Contains("cellKind", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SpecialCells_MissingSheetIsNotAnEmptySuccess()
    {
        var response = await _fixture.SendForFailureAsync("range.get-special-cells",
            new { sheetName = "MissingSheet", rangeAddress = "A1:A3", cellKind = "constants" });

        Assert.Equal("NotFound", response.ErrorCategory);
    }

    [Theory]
    [InlineData("constants")]
    [InlineData("visible")]
    public void SpecialCells_DisjointScopeDoesNotIncludeCellsBetweenAreas(string cellKind)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:C2",
            [[1, 99, 2], [3, 99, 4]]).Success);
        using var result = Discover(sheetName, "A1:A2,C1:C2", cellKind);

        Assert.Equal(4, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(["$A$1:$A$2", "$C$1:$C$2"], ReadAreas(result));
    }

    [Fact]
    public void SpecialCells_OverlappingAreasCountEachCellOnce()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A4",
            [[1], [2], [3], [4]]).Success);

        using var result = Discover(sheetName, "A1:A3,A2:A4", "constants");

        Assert.Equal(4, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(["$A$1:$A$4"], ReadAreas(result));
    }

    [Fact]
    public void SpecialCells_BlanksIncludeAllSidesOutsideUsedRange()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "C3", [[42]]).Success);

        using var result = Discover(sheetName, "B2:D4", "blanks");

        Assert.Equal(8, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(["$B$2:$D$2", "$B$3", "$D$3", "$B$4:$D$4"], ReadAreas(result));
        var value = _commands.GetValues(batch, sheetName, "C3");
        Assert.True(value.Success, value.ErrorMessage);
        Assert.Equal(42, Convert.ToInt32(value.Values[0][0], CultureInfo.InvariantCulture));
    }

    [Fact]
    public void SpecialCells_EmptySheetReturnsEntireRequestedBlankScope()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);

        using var result = Discover(sheetName, "B2:D4", "blanks");

        Assert.Equal(9, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(["$B$2:$D$4"], ReadAreas(result));
    }

    [Fact]
    public void SpecialCells_NamedRangeReturnsResolvedSheetAndScope()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var name = $"Discovery_{Guid.NewGuid():N}";
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A3",
            [[1], [null], [3]]).Success);
        _fixture.Send("namedrange.create", new
        {
            name,
            reference = $"'{sheetName}'!$A$1:$A$3"
        });
        _fixture.RegisterNamedRangeForCleanup(name);

        using var result = Discover(string.Empty, name, "constants");

        Assert.Equal(sheetName, result.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("$A$1:$A$3", result.RootElement.GetProperty("rangeAddress").GetString());
        Assert.Equal(2, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(["$A$1", "$A$3"], ReadAreas(result));
    }

    [Fact]
    public void SpecialCells_ErrorsIncludeStoredErrorsAndFormulaResults()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetFormulas(batch, sheetName, "A1", [["=1/0"]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range["C1"];
                cell.Formula = "=NA()";
                cell.Copy();
                cell.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                Assert.False(Convert.ToBoolean(cell.HasFormula, CultureInfo.InvariantCulture));
            }
            finally
            {
                context.App.CutCopyMode = 0;
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        using var result = Discover(sheetName, "A1:C1", "errors");

        Assert.Equal(2, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(["$A$1", "$C$1"], ReadAreas(result));
        using var constants = Discover(sheetName, "A1:C1", "constants");
        Assert.Equal(1, constants.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(["$C$1"], ReadAreas(constants));
    }

    [Fact]
    public void SpecialCells_VisibleExcludesFilteredRows()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B4",
            [["Key", "Value"], [1, "hidden"], [2, "visible"], [3, "visible"]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                range = sheet.Range["A1:B4"];
                range.AutoFilter(1, ">=2");
                Assert.True(sheet.FilterMode);
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

        using var result = Discover(sheetName, "A1:B4", "visible");

        Assert.Equal(6, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(["$A$1:$B$1", "$A$3:$B$4"], ReadAreas(result));
    }

    private JsonDocument Discover(
        string sheetName, string rangeAddress, string cellKind)
    {
        var response = _fixture.Send("range.get-special-cells",
            new { sheetName, rangeAddress, cellKind });
        var document = JsonDocument.Parse(response.Result!);
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        Assert.False(document.RootElement.TryGetProperty("errorMessage", out _));
        return document;
    }

    private static string[] ReadAreas(JsonDocument result) =>
        result.RootElement.GetProperty("areas").EnumerateArray()
            .Select(area => Assert.IsType<string>(area.GetString())).ToArray();
}
