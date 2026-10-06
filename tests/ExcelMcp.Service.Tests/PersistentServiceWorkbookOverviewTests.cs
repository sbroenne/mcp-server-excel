using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceWorkbookOverviewTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void Inspect_ReturnsBoundedMetadataAndPreview()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var tableName = $"OverviewTable_{Guid.NewGuid():N}";
        var definedName = $"OverviewName_{Guid.NewGuid():N}";
        var secondDefinedName = $"OverviewName_{Guid.NewGuid():N}";
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:C3",
        [
            [1, "a preview value longer than the limit", "ignored"],
            ["second row", "value", "ignored"],
            ["third row", "value", "ignored"]
        ]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1", [["=40+2"]],
            overwritePolicy: Sbroenne.ExcelMcp.Core.Commands.Range.OverwritePolicy.Allow).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "D1:E2",
        [
            ["Header A", "Header B"],
            [1, 2]
        ]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "XFD1048576", [["far edge"]]).Success);
        _fixture.Send("table.create", new
        {
            sheetName,
            tableName,
            rangeAddress = "D1:E2",
            hasHeaders = true
        });
        _fixture.Send("namedrange.create", new
        {
            name = definedName,
            reference = $"'{sheetName}'!$A$1"
        });
        _fixture.Send("namedrange.create", new
        {
            name = secondDefinedName,
            reference = $"'{sheetName}'!$B$1"
        });
        _fixture.RegisterNamedRangeForCleanup(definedName);
        _fixture.RegisterNamedRangeForCleanup(secondDefinedName);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                sheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
        var beforeView = ReadViewState();

        var response = _fixture.Send("workbook.inspect", new
        {
            sheetName,
            includePreview = true,
            maxItems = 1,
            maxPreviewRows = 1,
            maxPreviewColumns = 2,
            maxCellCharacters = 8,
            maxPreviewCharacters = 20
        });

        using var result = JsonDocument.Parse(response.Result!);
        var root = result.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());

        var sheets = root.GetProperty("sheets");
        Assert.Equal(1, sheets.GetProperty("count").GetInt32());
        Assert.Single(sheets.GetProperty("items").EnumerateArray());
        Assert.Equal(sheetName, sheets.GetProperty("items")[0].GetProperty("name").GetString());
        Assert.Equal("Hidden", sheets.GetProperty("items")[0].GetProperty("visibility").GetString());
        Assert.Equal(0, sheets.GetProperty("omittedCount").GetInt32());
        Assert.Equal(1_048_576, sheets.GetProperty("items")[0].GetProperty("usedRowCount").GetInt32());
        Assert.Equal(16_384, sheets.GetProperty("items")[0].GetProperty("usedColumnCount").GetInt32());

        var tables = root.GetProperty("tables");
        Assert.Equal(1, tables.GetProperty("count").GetInt32());
        Assert.Equal(tableName, tables.GetProperty("items")[0].GetProperty("name").GetString());
        Assert.Equal("$D$1:$E$2", tables.GetProperty("items")[0].GetProperty("rangeAddress").GetString());
        var returnedNames = root.GetProperty("definedNames").GetProperty("items")
            .EnumerateArray().Select(name => name.GetProperty("name").GetString()).ToArray();
        Assert.Contains(returnedNames, name => name == definedName || name == secondDefinedName);
        Assert.Equal(2, root.GetProperty("definedNames").GetProperty("count").GetInt32());
        Assert.Equal(1, root.GetProperty("definedNames").GetProperty("omittedCount").GetInt32());

        var preview = root.GetProperty("preview");
        Assert.Equal("$A$1:$B$1", preview.GetProperty("rangeAddress").GetString());
        Assert.Equal(1, preview.GetProperty("rowCount").GetInt32());
        Assert.Equal(2, preview.GetProperty("columnCount").GetInt32());
        Assert.Equal(1_048_575, preview.GetProperty("omittedRowCount").GetInt32());
        Assert.Equal(16_382, preview.GetProperty("omittedColumnCount").GetInt32());
        Assert.Equal(1, preview.GetProperty("values").GetArrayLength());
        Assert.Equal(2, preview.GetProperty("values")[0].GetArrayLength());
        Assert.Equal(1, preview.GetProperty("formulas").GetArrayLength());
        Assert.Equal(2, preview.GetProperty("formulas")[0].GetArrayLength());
        Assert.Equal(42d, preview.GetProperty("values")[0][0].GetDouble());
        Assert.Equal("=40+2", preview.GetProperty("formulas")[0][0].GetString());
        Assert.True(preview.GetProperty("textCharactersReturned").GetInt32() <= 20);
        Assert.True(preview.GetProperty("truncatedTextCellCount").GetInt32() > 0);
        Assert.Equal(beforeView, ReadViewState());
    }

    [Fact]
    public async Task Inspect_RequiresAWorksheetForPreviewAndRejectsUnboundedLimits()
    {
        var missingSheet = await _fixture.SendForFailureAsync("workbook.inspect", new
        {
            includePreview = true
        });
        Assert.Contains("sheetName is required", missingSheet.ErrorMessage, StringComparison.Ordinal);

        var excessiveRows = await _fixture.SendForFailureAsync("workbook.inspect", new
        {
            maxPreviewRows = 11
        });
        Assert.Contains("maxPreviewRows must be between 1 and 10", excessiveRows.ErrorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void Inspect_HandlesAnEmptySheetAndOnlyReturnsSelectedSections()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var response = _fixture.Send("workbook.inspect", new
        {
            sheetName,
            includeSheets = true,
            includeTables = false,
            includeDefinedNames = false,
            includePreview = true
        });

        using var result = JsonDocument.Parse(response.Result!);
        var root = result.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.True(root.TryGetProperty("sheets", out _));
        Assert.False(root.TryGetProperty("tables", out _));
        Assert.False(root.TryGetProperty("definedNames", out _));
        var sheet = Assert.Single(root.GetProperty("sheets").GetProperty("items").EnumerateArray());
        Assert.Equal("$A$1", sheet.GetProperty("usedRangeAddress").GetString());
        var preview = root.GetProperty("preview");
        Assert.Equal("$A$1", preview.GetProperty("rangeAddress").GetString());
        Assert.Equal(1, preview.GetProperty("rowCount").GetInt32());
        Assert.Equal(1, preview.GetProperty("columnCount").GetInt32());
        Assert.Equal(JsonValueKind.Null, preview.GetProperty("values")[0][0].ValueKind);
        Assert.Equal(string.Empty, preview.GetProperty("formulas")[0][0].GetString());
    }

    [Fact]
    public void Inspect_DistinguishesErrorResultsFromMatchingNumericValues()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A2",
            [[-2146826281d]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1",
            [["=1/0"]], overwritePolicy: Sbroenne.ExcelMcp.Core.Commands.Range.OverwritePolicy.Allow).Success);

        var response = _fixture.Send("workbook.inspect", new
        {
            sheetName,
            includeSheets = false,
            includeTables = false,
            includeDefinedNames = false,
            includePreview = true,
            rangeAddress = "A1:A2"
        });

        using var result = JsonDocument.Parse(response.Result!);
        var preview = result.RootElement.GetProperty("preview");
        Assert.Equal("#DIV/0!", preview.GetProperty("values")[0][0].GetString());
        Assert.Equal(-2146826281d, preview.GetProperty("values")[1][0].GetDouble());
        Assert.Equal("=1/0", preview.GetProperty("formulas")[0][0].GetString());
        Assert.Equal("-2146826281", preview.GetProperty("formulas")[1][0].GetString());
    }

    [Fact]
    public void Inspect_TruncatesTextWithoutSplittingSurrogatePairs()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A2",
            [["a😀b"], ["Z"]]).Success);

        var response = _fixture.Send("workbook.inspect", new
        {
            sheetName,
            includeSheets = false,
            includeTables = false,
            includeDefinedNames = false,
            includePreview = true,
            rangeAddress = "A1:A2",
            maxCellCharacters = 2,
            maxPreviewCharacters = 2
        });

        using var result = JsonDocument.Parse(response.Result!);
        var preview = result.RootElement.GetProperty("preview");
        Assert.Equal("a", preview.GetProperty("values")[0][0].GetString());
        Assert.Equal("Z", preview.GetProperty("values")[1][0].GetString());
        Assert.Equal(2, preview.GetProperty("textCharactersReturned").GetInt32());
        Assert.True(preview.GetProperty("truncatedTextCellCount").GetInt32() > 0);
    }

    private (string Sheet, string Address, bool Saved) ReadViewState()
    {
        return _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? activeSheet = null;
            Excel.Range? selection = null;
            try
            {
                activeSheet = (Excel.Worksheet)context.App.ActiveSheet;
                selection = (Excel.Range)context.App.Selection;
                return (activeSheet.Name, selection.Address, context.Book.Saved);
            }
            finally
            {
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref activeSheet);
            }
        });
    }
}
