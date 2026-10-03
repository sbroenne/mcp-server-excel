using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "CellStyles")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceCellStyleTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public async Task CellStyle_LifecycleUpdatesExistingUsersAndPersists()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string styleName = $"Style_{Guid.NewGuid():N}";
        var prepared = _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1"],
            formatOptions = new { bold = true, fillThemeColor = 5, indentLevel = 2, horizontalAlignment = "left", numberFormat = "0.00%" }
        });
        Assert.True(prepared.Success, prepared.ErrorMessage);
        var created = _fixture.Send("workbook.create-cell-style", new
        {
            styleName,
            sourceSheetName = sheetName,
            sourceCellAddress = "A1"
        });
        using (var result = JsonDocument.Parse(created.Result!))
        {
            Assert.False(result.RootElement.GetProperty("style").GetProperty("builtIn").GetBoolean());
            Assert.Equal(styleName, result.RootElement.GetProperty("style").GetProperty("name").GetString());
            Assert.True(result.RootElement.GetProperty("style").GetProperty("format").GetProperty("font").GetProperty("bold").GetBoolean());
        }
        Assert.True(_fixture.Send("rangeformat.set-style", new { sheetName, rangeAddress = "C1", styleName }).Success);
        var changed = _fixture.Send("workbook.update-cell-style", new
        {
            styleName,
            styleOptions = new { formatOptions = new { bold = false, fillThemeColor = 6, numberFormat = "0.000" }, includeFont = true }
        });
        Assert.True(changed.Success, changed.ErrorMessage);
        using (var read = ReadFormat(sheetName, "C1"))
        {
            Assert.False(read.RootElement.GetProperty("font").GetProperty("bold").GetBoolean());
            Assert.Equal(6, read.RootElement.GetProperty("fill").GetProperty("color").GetProperty("themeColor").GetInt32());
            Assert.Equal("0.000", read.RootElement.GetProperty("numberFormat").GetString());
        }
        using (var listed = JsonDocument.Parse(_fixture.Send("workbook.list-cell-styles", new { }).Result!))
            Assert.Contains(listed.RootElement.GetProperty("styles").EnumerateArray(),
                style => style.GetProperty("name").GetString() == styleName && !style.GetProperty("builtIn").GetBoolean());
        await _fixture.SaveAndReopenAsync();
        var readback = _fixture.Send("workbook.get-cell-style", new { styleName });
        using (var reopened = JsonDocument.Parse(readback.Result!))
            Assert.Equal("0.000", reopened.RootElement.GetProperty("style").GetProperty("format").GetProperty("numberFormat").GetString());
        Assert.True(_fixture.Send("workbook.delete-cell-style", new { styleName }).Success);
        var missing = await _fixture.SendForFailureAsync("workbook.get-cell-style", new { styleName });
        Assert.False(missing.Success);
        using var remaining = ReadFormat(sheetName, "C1");
        Assert.NotEqual(styleName, remaining.RootElement.GetProperty("styleName").GetString());
    }

    [Fact]
    public void GetStyle_ReportsNativeCustomStyleInsteadOfAssumingBuiltIn()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string styleName = $"Native_{Guid.NewGuid():N}";
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Styles? styles = null;
            Excel.Style? style = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                range = sheet.Range["A1"];
                styles = ctx.Book.Styles;
                style = styles.Add(styleName, range);
                range.Style = styleName;
            }
            finally
            {
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref styles);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var response = _fixture.Send("rangeformat.get-style", new { sheetName, rangeAddress = "A1" });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal(styleName, result.RootElement.GetProperty("styleName").GetString());
        Assert.False(result.RootElement.GetProperty("isBuiltInStyle").GetBoolean());
    }

    private JsonDocument ReadFormat(string sheetName, string rangeAddress)
    {
        var response = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        return JsonDocument.Parse(result.RootElement.GetProperty("cells")[0].GetProperty("stored").GetRawText());
    }

    [Theory]
    [InlineData("Left", "xlEdgeLeft")]
    [InlineData("Top", "xlEdgeTop")]
    [InlineData("Bottom", "xlEdgeBottom")]
    [InlineData("Right", "xlEdgeRight")]
    [InlineData("DiagonalDown", "xlDiagonalDown")]
    [InlineData("DiagonalUp", "xlDiagonalUp")]
    public void CellStyle_SelectedBordersAreAppliedToExistingUsers(string position, string edge)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string styleName = $"Border_{Guid.NewGuid():N}";
        Assert.True(_fixture.Send("workbook.create-cell-style", new
        {
            styleName,
            sourceSheetName = sheetName,
            sourceCellAddress = "A1"
        }).Success);
        Assert.True(_fixture.Send("rangeformat.set-style", new { sheetName, rangeAddress = "C1", styleName }).Success);
        var changed = _fixture.Send("workbook.update-cell-style", new
        {
            styleName,
            styleOptions = new
            {
                includeBorder = true,
                formatOptions = new
                {
                    borders = new[]
            {
                new { position, lineStyle = "dash", color = "#123456" }
            }
                }
            }
        });
        Assert.True(changed.Success, changed.ErrorMessage);
        using (var definition = JsonDocument.Parse(changed.Result!))
        {
            var selected = Assert.Single(definition.RootElement.GetProperty("style").GetProperty("format").GetProperty("borders").EnumerateArray(),
                item => item.GetProperty("edge").GetString() == edge);
            Assert.Equal("#123456", selected.GetProperty("color").GetProperty("rgb").GetString());
        }
        using var read = ReadFormat(sheetName, "C1");
        var border = Assert.Single(read.RootElement.GetProperty("borders").EnumerateArray(),
            item => item.GetProperty("edge").GetString() == edge);
        Assert.Equal(-4115, border.GetProperty("lineStyle").GetInt32());
        Assert.Equal("#123456", border.GetProperty("color").GetProperty("rgb").GetString());
    }

    [Theory]
    [InlineData("update-cell-style")]
    [InlineData("delete-cell-style")]
    public async Task CellStyle_BuiltInMutationsAreRejectedWithoutChangingDefinition(string action)
    {
        var before = _fixture.Send("workbook.get-cell-style", new { styleName = "Normal" });
        object args = action == "update-cell-style" ? new
        {
            styleName = "Normal",
            styleOptions = new { formatOptions = new { bold = true } }
        } : new { styleName = "Normal" };
        var rejected = await _fixture.SendForFailureAsync($"workbook.{action}", args);
        Assert.False(rejected.Success);
        Assert.Contains("Built-in", rejected.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(before.Result, _fixture.Send("workbook.get-cell-style", new { styleName = "Normal" }).Result);
    }

    [Theory]
    [InlineData("includeFont")]
    [InlineData("includeNumber")]
    [InlineData("includeAlignment")]
    [InlineData("includeBorder")]
    [InlineData("includePatterns")]
    [InlineData("includeProtection")]
    public void CellStyle_ExplicitFalseInclusionFlagsArePreservedWhenDefinitionChanges(string flag)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string styleName = $"Flags_{Guid.NewGuid():N}";
        Assert.True(_fixture.Send("workbook.create-cell-style", new
        {
            styleName,
            sourceSheetName = sheetName,
            sourceCellAddress = "A1"
        }).Success);
        Assert.True(_fixture.Send("workbook.update-cell-style", new
        {
            styleName,
            styleOptions = new Dictionary<string, object?> { [flag] = false }
        }).Success);
        var changed = _fixture.Send("workbook.update-cell-style", new
        {
            styleName,
            styleOptions = new
            {
                formatOptions = new
                {
                    bold = true,
                    fillColor = "#123456",
                    numberFormat = "0.00",
                    horizontalAlignment = "left"
                }
            }
        });
        Assert.True(changed.Success, changed.ErrorMessage);
        using var read = JsonDocument.Parse(changed.Result!);
        Assert.False(read.RootElement.GetProperty("style").GetProperty(flag).GetBoolean());
    }

    [Fact]
    public async Task CellStyle_InvalidInsideBorderPreservesDefinition()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string styleName = $"Inside_{Guid.NewGuid():N}";
        Assert.True(_fixture.Send("workbook.create-cell-style", new
        {
            styleName,
            sourceSheetName = sheetName,
            sourceCellAddress = "A1"
        }).Success);
        var before = _fixture.Send("workbook.get-cell-style", new { styleName });
        var rejected = await _fixture.SendForFailureAsync("workbook.update-cell-style", new
        {
            styleName,
            styleOptions = new
            {
                formatOptions = new
                {
                    bold = true,
                    borders = new[] { new { position = "InsideHorizontal", lineStyle = "dash" } }
                }
            }
        });
        Assert.False(rejected.Success);
        Assert.Equal(before.Result, _fixture.Send("workbook.get-cell-style", new { styleName }).Result);
    }

    [Fact]
    public async Task CellStyle_DuplicateNameAndMultiCellSourceDoNotChangeCatalogue()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var before = _fixture.Send("workbook.list-cell-styles", new { });
        var duplicate = await _fixture.SendForFailureAsync("workbook.create-cell-style", new
        {
            styleName = "Normal",
            sourceSheetName = sheetName,
            sourceCellAddress = "A1"
        });
        Assert.False(duplicate.Success);
        var multiple = await _fixture.SendForFailureAsync("workbook.create-cell-style", new
        {
            styleName = $"Invalid_{Guid.NewGuid():N}",
            sourceSheetName = sheetName,
            sourceCellAddress = "A1:B2"
        });
        Assert.False(multiple.Success);
        Assert.Equal(before.Result, _fixture.Send("workbook.list-cell-styles", new { }).Result);
    }

    [Theory]
    [InlineData("get-cell-style")]
    [InlineData("update-cell-style")]
    [InlineData("delete-cell-style")]
    public async Task CellStyle_MissingNamesFailWithoutChangingCatalogue(string action)
    {
        var before = _fixture.Send("workbook.list-cell-styles", new { });
        string styleName = $"Missing_{Guid.NewGuid():N}";
        object args = action == "update-cell-style"
            ? new { styleName, styleOptions = new { includeFont = false } }
            : new { styleName };
        var response = await _fixture.SendForFailureAsync($"workbook.{action}", args);
        Assert.Equal("NotFound", response.ErrorCategory);
        Assert.Equal(before.Result, _fixture.Send("workbook.list-cell-styles", new { }).Result);
    }

    [Fact]
    public void CellStyle_InspectionReturnsSixNativeBordersAndNullableUnsetFields()
    {
        using var result = JsonDocument.Parse(_fixture.Send("workbook.get-cell-style", new { styleName = "Normal" }).Result!);
        var style = result.RootElement.GetProperty("style");
        Assert.True(style.GetProperty("builtIn").GetBoolean());
        Assert.Equal("Normal", style.GetProperty("name").GetString());
        Assert.Equal(6, style.GetProperty("format").GetProperty("borders").GetArrayLength());
        Assert.NotEmpty(result.RootElement.GetProperty("readLimitations").EnumerateArray());
    }

    [Fact]
    public void CellStyle_CapturesInactiveSourceWithoutChangingSelection()
    {
        var source = _fixture.CreateTestSheet(_fixture.BatchToken);
        var active = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_fixture.Send("rangeformat.format", new
        {
            sheetName = source,
            rangeAddresses = (string[])["AA10:AB11"],
            formatOptions = new
            {
                fillThemeColor = 6,
                fontThemeColor = 5,
                indentLevel = 2,
                horizontalAlignment = "left",
                borders = new[] { new { position = "DiagonalUp", lineStyle = "dash", color = "#123456" } }
            }
        }).Success);
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? selection = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[active];
                sheet.Activate();
                selection = sheet.Range["D7"];
                selection.Select();
            }
            finally
            {
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var captured = _fixture.Send("workbook.create-cell-style", new
        {
            styleName = $"Inactive_{Guid.NewGuid():N}",
            sourceSheetName = source,
            sourceCellAddress = "AA10"
        });
        Assert.True(captured.Success, captured.ErrorMessage);
        using var result = JsonDocument.Parse(captured.Result!);
        Assert.Equal(6, result.RootElement.GetProperty("style").GetProperty("format").GetProperty("fill").GetProperty("color").GetProperty("themeColor").GetInt32());
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? selection = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.ActiveSheet;
                Assert.Equal(active, sheet.Name);
                selection = (Excel.Range)ctx.App.Selection;
                Assert.Equal("$D$7", selection.Address);
            }
            finally
            {
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    [Theory]
    [InlineData((int)Excel.XlSheetVisibility.xlSheetHidden)]
    [InlineData((int)Excel.XlSheetVisibility.xlSheetVeryHidden)]
    public async Task CellStyle_HiddenSourceFailsWithoutChangingVisibilityOrCatalogue(int visibility)
    {
        var source = _fixture.CreateTestSheet(_fixture.BatchToken);
        var active = _fixture.CreateTestSheet(_fixture.BatchToken);
        var before = _fixture.Send("workbook.list-cell-styles", new { });
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[source];
                sheet.Visible = (Excel.XlSheetVisibility)visibility;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var rejected = await _fixture.SendForFailureAsync("workbook.create-cell-style", new
        {
            styleName = $"Hidden_{Guid.NewGuid():N}",
            sourceSheetName = source,
            sourceCellAddress = "A1"
        });
        Assert.False(rejected.Success);
        Assert.Equal("InvalidInput", rejected.ErrorCategory);
        Assert.Contains("visible", rejected.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(before.Result, _fixture.Send("workbook.list-cell-styles", new { }).Result);
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Worksheet? activeSheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[source];
                Assert.Equal((Excel.XlSheetVisibility)visibility, sheet.Visible);
                activeSheet = (Excel.Worksheet)ctx.Book.ActiveSheet;
                Assert.Equal(active, activeSheet.Name);
            }
            finally
            {
                ComUtilities.Release(ref activeSheet);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }
}
