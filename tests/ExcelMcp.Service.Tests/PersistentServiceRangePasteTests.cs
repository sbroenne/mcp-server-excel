using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceRangePasteTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void Copy_FormatsPreserveOccupiedValuesAndFormulasAcrossSheets()
    {
        var source = _fixture.CreateTestSheet(_fixture.BatchToken);
        var target = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, source, "A1:B1", [[1, 2]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, target, "D1", [["keep"]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, target, "E1", [["=42"]]).Success);
        Assert.True(_commands.SetNumberFormat(_fixture.BatchToken, source, "A1:B1", "0.000").Success);
        _fixture.Send("rangeformat.format", new
        {
            sheetName = source,
            rangeAddresses = (string[])["A1:B1"],
            formatOptions = new
            {
                bold = true,
                fillColor = "#123456"
            }
        });
        var response = _fixture.Send("range.copy", new
        {
            sourceSheet = source,
            sourceRange = "A1:B1",
            targetSheet = target,
            targetRange = "D1",
            pasteKind = "formats"
        });
        using var copy = JsonDocument.Parse(response.Result!);
        Assert.Equal("$D$1:$E$1", copy.RootElement.GetProperty("destinationAddress").GetString());
        var content = _commands.GetFormulas(_fixture.BatchToken, target, "D1:E1");
        Assert.True(content.Success);
        Assert.Equal("keep", content.Values[0][0]);
        Assert.Equal("=42", content.Formulas[0][1]);
        var read = _fixture.Send("rangeformat.get-format", new
        {
            sheetName = target,
            rangeAddress = "D1:E1"
        });
        using var formats = JsonDocument.Parse(read.Result!);
        Assert.All(formats.RootElement.GetProperty("cells").EnumerateArray(), cell =>
        {
            var snapshot = cell.GetProperty("stored");
            Assert.True(snapshot.GetProperty("font").GetProperty("bold").GetBoolean());
            Assert.Equal("#123456", snapshot.GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
            Assert.Equal("0.000", snapshot.GetProperty("numberFormat").GetString());
        });
        AssertClipboardReleased();
    }

    [Fact]
    public void Copy_ValidationPreservesContentAndVisualFormatting()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "B1", [[5]]).Success);
        Assert.True(_commands.SetNumberFormat(_fixture.BatchToken, sheetName, "B1", "0.000").Success);
        _fixture.Send("rangeformat.validate-range", new
        {
            sheetName,
            rangeAddress = "A1",
            validationType = "whole",
            validationOperator = "between",
            formula1 = "1",
            formula2 = "10"
        });
        _fixture.Send("range.copy", new
        {
            sourceSheet = sheetName,
            sourceRange = "A1",
            targetSheet = sheetName,
            targetRange = "B1",
            pasteKind = "validation"
        });
        var values = _commands.GetValues(_fixture.BatchToken, sheetName, "B1");
        Assert.True(values.Success);
        Assert.Equal(5, Convert.ToDouble(values.Values[0][0], CultureInfo.InvariantCulture));
        var formats = _commands.GetNumberFormats(_fixture.BatchToken, sheetName, "B1");
        Assert.True(formats.Success);
        Assert.Equal("0.000", formats.Formats[0][0]);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            Excel.Validation? validation = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range["B1"];
                validation = cell.Validation;
                Assert.Equal((int)Excel.XlDVType.xlValidateWholeNumber, validation.Type);
                Assert.Equal("1", validation.Formula1);
                Assert.Equal("10", validation.Formula2);
            }
            finally
            {
                ComUtilities.Release(ref validation);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
        AssertClipboardReleased();
    }

    [Fact]
    public async Task Copy_TransposeChecksActualExpandedDestinationBeforeMutation()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:C2",
            [[1, 2, 3], [4, 5, 6]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "E3", [["keep"]]).Success);
        var arguments = new
        {
            sourceSheet = sheetName,
            sourceRange = "A1:C2",
            targetSheet = sheetName,
            targetRange = "D1",
            pasteKind = "values",
            transpose = true
        };
        var rejected = await _fixture.SendForFailureAsync("range.copy", arguments);
        Assert.Equal("Conflict", rejected.ErrorCategory);
        Assert.Contains("$E$3", rejected.ErrorMessage);
        Assert.True(_commands.ClearContents(_fixture.BatchToken, sheetName, "E3").Success);
        var response = _fixture.Send("range.copy", arguments);
        using var copied = JsonDocument.Parse(response.Result!);
        Assert.Equal("$D$1:$E$3", copied.RootElement.GetProperty("destinationAddress").GetString());
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "D1:E3");
        Assert.True(read.Success);
        Assert.Equal([1d, 4d], read.Values[0].Select(value => Convert.ToDouble(value, CultureInfo.InvariantCulture)));
        Assert.Equal([2d, 5d], read.Values[1].Select(value => Convert.ToDouble(value, CultureInfo.InvariantCulture)));
        Assert.Equal([3d, 6d], read.Values[2].Select(value => Convert.ToDouble(value, CultureInfo.InvariantCulture)));
        AssertClipboardReleased();
    }

    [Fact]
    public void Copy_SkipBlanksDoesNotRequirePermissionToReplaceUntouchedCells()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:B1", [[42, null]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "E1", [["keep"]]).Success);
        _fixture.Send("range.copy", new
        {
            sourceSheet = sheetName,
            sourceRange = "A1:B1",
            targetSheet = sheetName,
            targetRange = "D1",
            pasteKind = "values",
            skipBlanks = true
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "D1:E1");
        Assert.True(read.Success);
        Assert.Equal(42, Convert.ToDouble(read.Values[0][0], CultureInfo.InvariantCulture));
        Assert.Equal("keep", read.Values[0][1]);
        AssertClipboardReleased();
    }

    private void AssertClipboardReleased()
    {
        _fixture.ExecuteRawVerification((context, _) =>
            Assert.Equal(0, Convert.ToInt32(context.App.CutCopyMode, CultureInfo.InvariantCulture)));
    }

    public static TheoryData<string, bool, bool> ContentPasteOptions
    {
        get
        {
            TheoryData<string, bool, bool> data = [];
            foreach (var kind in new[] { "all", "values", "formulas" })
                foreach (var transpose in new[] { false, true })
                    foreach (var skipBlanks in new[] { false, true })
                        data.Add(kind, transpose, skipBlanks);
            return data;
        }
    }

    [Theory]
    [MemberData(nameof(ContentPasteOptions))]
    public void Copy_ContentKindsHonorNativeTransposeAndSkipBlanks(string pasteKind, bool transpose, bool skipBlanks)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:B2", [[1, null], [3, 4]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "D1:E2", [[99, 99], [99, 99]]).Success);
        _fixture.Send("range.copy", new
        {
            sourceSheet = sheetName,
            sourceRange = "A1:B2",
            targetSheet = sheetName,
            targetRange = "D1",
            pasteKind,
            transpose,
            skipBlanks,
            overwritePolicy = "allow"
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "D1:E2");
        Assert.True(read.Success);
        double? blank = skipBlanks ? 99 : null;
        double?[][] expected = transpose ? [[1, 3], [blank, 4]] : [[1, blank], [3, 4]];
        for (int row = 0; row < 2; row++)
            Assert.Equal(expected[row], read.Values[row].Select(value =>
                value is null ? (double?)null : Convert.ToDouble(value, CultureInfo.InvariantCulture)));
        AssertClipboardReleased();
    }

    [Fact]
    public void Copy_FormulasAdjustReferencesAndDoNotReplaceDestinationNumberFormats()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[5]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A2", [["=A1*2"]]).Success);
        Assert.True(_commands.SetNumberFormat(_fixture.BatchToken, sheetName, "A2", "0.000").Success);
        Assert.True(_commands.SetNumberFormat(_fixture.BatchToken, sheetName, "D2", "0%").Success);
        _fixture.Send("range.copy", new
        {
            sourceSheet = sheetName,
            sourceRange = "A1:A2",
            targetSheet = sheetName,
            targetRange = "D1",
            pasteKind = "formulas"
        });
        var formulas = _commands.GetFormulas(_fixture.BatchToken, sheetName, "D1:D2");
        Assert.True(formulas.Success);
        Assert.Equal(5, Convert.ToDouble(formulas.Values[0][0], CultureInfo.InvariantCulture));
        Assert.Equal("=D1*2", formulas.Formulas[1][0]);
        Assert.Equal(10, Convert.ToDouble(formulas.Values[1][0], CultureInfo.InvariantCulture));
        var formats = _commands.GetNumberFormats(_fixture.BatchToken, sheetName, "D2");
        Assert.True(formats.Success);
        Assert.Equal("0%", formats.Formats[0][0]);
    }

    [Fact]
    public async Task Copy_ProtectedTargetFailureStillClearsCopyMode()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[5]]).Success);
        void Protect(bool enabled) => _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                if (enabled)
                    sheet.Protect();
                else
                    sheet.Unprotect();
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
        Protect(true);
        try
        {
            var response = await _fixture.SendForFailureAsync("range.copy", new
            {
                sourceSheet = sheetName,
                sourceRange = "A1",
                targetSheet = sheetName,
                targetRange = "D1",
                pasteKind = "values"
            });
            Assert.False(response.Success);
            Assert.False(string.IsNullOrWhiteSpace(response.ErrorMessage));
            AssertClipboardReleased();
            var read = _commands.GetValues(_fixture.BatchToken, sheetName, "D1");
            Assert.True(read.Success);
            Assert.Null(read.Values[0][0]);
        }
        finally
        {
            Protect(false);
        }
    }

    [Fact]
    public void Copy_FormatsTransferConditionalRulesAndProtectionWithoutContent()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[1]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "D1", [[99]]).Success);
        _fixture.Send("conditionalformat.add-rule", new
        {
            sheetName,
            rangeAddress = "A1",
            ruleType = "expression",
            formula1 = "=A1>0",
            interiorColor = "#FF0000"
        });
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range["A1"];
                cell.Locked = false;
                cell.FormulaHidden = true;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
        _fixture.Send("range.copy", new
        {
            sourceSheet = sheetName,
            sourceRange = "A1",
            targetSheet = sheetName,
            targetRange = "D1",
            pasteKind = "formats"
        });
        var formatResponse = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress = "D1", view = "both" });
        using var format = JsonDocument.Parse(formatResponse.Result!);
        var cellRead = format.RootElement.GetProperty("cells")[0];
        Assert.False(cellRead.GetProperty("stored").GetProperty("locked").GetBoolean());
        Assert.True(cellRead.GetProperty("stored").GetProperty("formulaHidden").GetBoolean());
        Assert.Equal("#FF0000", cellRead.GetProperty("displayed").GetProperty("fill")
            .GetProperty("color").GetProperty("rgb").GetString());
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "D1");
        Assert.True(read.Success);
        Assert.Equal(99, Convert.ToDouble(read.Values[0][0], CultureInfo.InvariantCulture));
    }
}
