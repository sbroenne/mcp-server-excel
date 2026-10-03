using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "NativeDataCleanup")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceNativeCleanupTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void RemoveDuplicates_CountsBlankRecordsAndPreservesCellsBelowTheScope()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B7",
            [["Key", "Amount"], [1, 10], [1, 20], [null, null], [null, null], [2, 30], ["below", "keep"]]).Success);
        int[] keyColumns = [1];
        var response = _fixture.Send("rangeedit.remove-duplicates", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B6",
            keyColumns,
            hasHeaders = true
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(2, result.RootElement.GetProperty("removedRows").GetInt32());
        Assert.Equal(3, result.RootElement.GetProperty("remainingRows").GetInt32());
        Assert.Equal("$A$1:$B$4", result.RootElement.GetProperty("remainingRange").GetString());
        var values = _commands.GetValues(_fixture.BatchToken, sheet, "A1:B7");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal(1d, Convert.ToDouble(values.Values[1][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Null(values.Values[2][0]);
        Assert.Equal(2d, Convert.ToDouble(values.Values[3][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal("below", values.Values[6][0]);
        Assert.Equal("keep", values.Values[6][1]);
    }

    [Fact]
    public async Task TextToColumns_ProtectsTrailingEmptyOutputBeforeWriting()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:A2",
            [["x,,"], ["\"x,y\",z"]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "F1", [["protected"]]).Success);
        var rejected = await _fixture.SendForFailureAsync("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1:A2",
            destinationCell = "D1",
            options = new { comma = true }
        });
        Assert.False(rejected.Success);
        Assert.Contains("$F$1", rejected.ErrorMessage);
        var unchanged = _commands.GetValues(_fixture.BatchToken, sheet, "D1:F2");
        Assert.True(unchanged.Success, unchanged.ErrorMessage);
        Assert.Null(unchanged.Values[0][0]);
        Assert.Equal("protected", unchanged.Values[0][2]);
        var response = _fixture.Send("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1:A2",
            destinationCell = "D1",
            options = new { comma = true },
            overwritePolicy = "allow"
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("$D$1:$F$2", result.RootElement.GetProperty("destinationRange").GetString());
        var values = _commands.GetValues(_fixture.BatchToken, sheet, "D1:F2");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal("x", values.Values[0][0]);
        Assert.Equal("x,y", values.Values[1][0]);
        Assert.Equal("z", values.Values[1][1]);
        Assert.True(values.Values[0][2] is null or "");
    }

    [Theory]
    [InlineData("a,b,c", 3, "F2")]
    [InlineData("a,b,,", 4, "G2")]
    public async Task TextToColumns_LaterWiderRowsProtectTheCompleteDestination(
        string laterRow, int width, string protectedCell)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:A2", [["x"], [laterRow]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, protectedCell, [["protected"]]).Success);
        var before = ReadNativeView();
        var rejected = await _fixture.SendForFailureAsync("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1:A2",
            destinationCell = "D1",
            options = new { comma = true }
        });
        Assert.False(rejected.Success);
        Assert.Equal("Conflict", rejected.ErrorCategory);
        Assert.Contains("$" + protectedCell[0] + "$2", rejected.ErrorMessage);
        Assert.Equal(before, ReadNativeView());
        var source = _commands.GetValues(_fixture.BatchToken, sheet, "A1:A2");
        Assert.True(source.Success, source.ErrorMessage);
        Assert.Equal("x", source.Values[0][0]);
        Assert.Equal(laterRow, source.Values[1][0]);
        var untouched = _commands.GetValues(_fixture.BatchToken, sheet, "D1:G2");
        Assert.True(untouched.Success, untouched.ErrorMessage);
        for (int row = 0; row < 2; row++)
        {
            for (int column = 0; column < 4; column++)
            {
                Assert.Equal(row == 1 && column == width - 1 ? "protected" : null,
                    untouched.Values[row][column]);
            }
        }
        var response = _fixture.Send("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1:A2",
            destinationCell = "D1",
            options = new { comma = true },
            overwritePolicy = "allow"
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal(width, result.RootElement.GetProperty("outputColumns").GetInt32());
        Assert.Equal($"$D$1:${protectedCell[0]}$2", result.RootElement.GetProperty("destinationRange").GetString());
        var parsed = _commands.GetValues(_fixture.BatchToken, sheet, "D1:H2");
        Assert.True(parsed.Success, parsed.ErrorMessage);
        Assert.Equal("x", parsed.Values[0][0]);
        Assert.Equal("a", parsed.Values[1][0]);
        Assert.Equal("b", parsed.Values[1][1]);
        Assert.True(width == 3 ? Equals("c", parsed.Values[1][2]) : parsed.Values[1][2] is null or "");
        Assert.True(width == 3 || parsed.Values[1][3] is null or "");
        Assert.Null(parsed.Values[1][width]);
    }

    [Theory]
    [InlineData("x", "x,y,z", 3)]
    [InlineData("x,,", "y,,", 3)]
    [InlineData("\"x,y\",z", "\"a,b\",c", 2)]
    [InlineData("x", "y", 1)]
    public void TextToColumns_UsesTheCompleteNativeOutputWidth(string first, string second, int width)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:A2", [[first], [second]]).Success);
        var response = _fixture.Send("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1:A2",
            destinationCell = "D1",
            options = new { comma = true }
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(width, result.RootElement.GetProperty("outputColumns").GetInt32());
        var read = _commands.GetValues(_fixture.BatchToken, sheet, "D1:G2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Null(read.Values[0][width]);
        Assert.Null(read.Values[1][width]);
    }

    [Fact]
    public void TextToColumns_InPlaceParsingPreservesExplicitTextFieldsAndSkipsFields()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:A2",
            [["001,ignore,10"], ["002,ignore,20"]]).Success);
        var response = _fixture.Send("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1:A2",
            destinationCell = "A1",
            options = new
            {
                comma = true,
                fields = new[]
                {
                    new { position = 1, dataType = "Text" },
                    new { position = 2, dataType = "Skip" },
                    new { position = 3, dataType = "General" }
                }
            }
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal(2, result.RootElement.GetProperty("outputColumns").GetInt32());
        var values = _commands.GetValues(_fixture.BatchToken, sheet, "A1:C2");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal("001", values.Values[0][0]);
        Assert.Equal("002", values.Values[1][0]);
        Assert.Equal(10, Convert.ToDouble(values.Values[0][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Null(values.Values[0][2]);
    }

    [Fact]
    public void TextToColumns_FixedWidthUsesZeroBasedPositions()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:A2",
            [["001ABC10"], ["002DEF20"]]).Success);
        _fixture.Send("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1:A2",
            destinationCell = "D1",
            options = new
            {
                mode = "FixedWidth",
                fields = new[]
                {
                    new { position = 0, dataType = "Text" },
                    new { position = 3, dataType = "Skip" },
                    new { position = 6, dataType = "General" }
                }
            }
        });
        var values = _commands.GetValues(_fixture.BatchToken, sheet, "D1:F2");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal("001", values.Values[0][0]);
        Assert.Equal("002", values.Values[1][0]);
        Assert.Equal(10, Convert.ToDouble(values.Values[0][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(20, Convert.ToDouble(values.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Null(values.Values[0][2]);
    }

    [Fact]
    public void TextToColumns_CustomSeparatorsAndDateConversionUseNativeExcel()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1", [["1.234,5-;2026-02-03"]]).Success);
        _fixture.Send("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1",
            destinationCell = "D1",
            options = new
            {
                semicolon = true,
                decimalSeparator = ",",
                thousandsSeparator = ".",
                fields = new[] { new { position = 2, dataType = "Ymd" } }
            }
        });
        var values = _commands.GetValues(_fixture.BatchToken, sheet, "D1:E1");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal(-1234.5, Convert.ToDouble(values.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(new DateTime(2026, 2, 3).ToOADate(),
            Convert.ToDouble(values.Values[0][1], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemoveDuplicates_UsesCompositeKeysAndExplicitHeaders(bool hasHeaders)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        List<List<object?>> values = [[1, "a", "first"], [1, "b", "second"], [1, "a", "duplicate"]];
        if (hasHeaders)
            values.Insert(0, ["One", "Two", "Keep"]);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, $"A1:C{values.Count}", values).Success);
        int[] keyColumns = [1, 2];
        var response = _fixture.Send("rangeedit.remove-duplicates", new
        {
            sheetName = sheet,
            rangeAddress = $"A1:C{values.Count}",
            keyColumns,
            hasHeaders
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal(1, result.RootElement.GetProperty("removedRows").GetInt32());
        Assert.Equal(2, result.RootElement.GetProperty("remainingRows").GetInt32());
        var read = _commands.GetValues(_fixture.BatchToken, sheet, $"A1:C{values.Count}");
        Assert.True(read.Success, read.ErrorMessage);
        int start = hasHeaders ? 1 : 0;
        Assert.Equal("first", read.Values[start][2]);
        Assert.Equal("second", read.Values[start + 1][2]);
        Assert.Null(read.Values[start + 2][0]);
    }

    [Fact]
    public void RemoveDuplicates_AllBlankRecordsStillHaveExactCounts()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        int[] keyColumns = [1];
        var response = _fixture.Send("rangeedit.remove-duplicates", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B3",
            keyColumns,
            hasHeaders = false
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal(2, result.RootElement.GetProperty("removedRows").GetInt32());
        Assert.Equal(1, result.RootElement.GetProperty("remainingRows").GetInt32());
        Assert.Equal("$A$1:$B$1", result.RootElement.GetProperty("remainingRange").GetString());
    }

    [Fact]
    public async Task TextToColumns_FailedPreflightClosesScratchAndRestoresTheView()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1", [["x,y,z"]]).Success);
        var before = ReadNativeView();
        var response = await _fixture.SendForFailureAsync("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1",
            destinationCell = "XFD1",
            options = new { comma = true }
        });
        Assert.False(response.Success);
        Assert.Contains("boundaries", response.ErrorMessage);
        Assert.Equal(before, ReadNativeView());
        var source = _commands.GetValues(_fixture.BatchToken, sheet, "A1");
        Assert.True(source.Success, source.ErrorMessage);
        Assert.Equal("x,y,z", source.Values[0][0]);
    }

    [Theory]
    [InlineData(false, 3)]
    [InlineData(true, 2)]
    public void TextToColumns_ConsecutiveCustomDelimitersAndSingleQuotes(bool consecutive, int width)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1", [["''a|b'||c"]]).Success);
        var before = ReadNativeView();
        var response = _fixture.Send("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1",
            destinationCell = "D1",
            options = new
            {
                otherDelimiter = "|",
                qualifier = "SingleQuote",
                consecutiveDelimiters = consecutive
            }
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal(width, result.RootElement.GetProperty("outputColumns").GetInt32());
        var read = _commands.GetValues(_fixture.BatchToken, sheet, "D1:F1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal("a|b", read.Values[0][0]);
        Assert.Equal("c", read.Values[0][width - 1]);
        Assert.Equal(before, ReadNativeView());
    }

    [Fact]
    public void TextToColumns_UsesNativeFormulaTextWithoutChangingTheSourceFormula()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheet, "A1", [["=\"x,y\""]]).Success);
        _fixture.Send("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1",
            destinationCell = "D1",
            options = new { comma = true }
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheet, "D1:E1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal("=\"x", read.Values[0][0]);
        Assert.Equal("y\"", read.Values[0][1]);
        var formulas = _commands.GetFormulas(_fixture.BatchToken, sheet, "A1");
        Assert.True(formulas.Success, formulas.ErrorMessage);
        Assert.Equal("=\"x,y\"", formulas.Formulas[0][0]);
    }

    [Theory]
    [InlineData("0", "0.0")]
    [InlineData("0.00", "0.00")]
    public void RemoveDuplicates_PreviewMatchesNativeNumberFormats(string firstFormat, string secondFormat)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B3",
            [[1.01, "first"], [1.02, "second"], [1.01, "third"]]).Success);
        _fixture.Send("range.set-number-format", new { sheetName = sheet, rangeAddress = "A1", formatCode = firstFormat });
        _fixture.Send("range.set-number-format", new { sheetName = sheet, rangeAddress = "A2:A3", formatCode = secondFormat });
        int[] keyColumns = [1];
        var response = _fixture.Send("rangeedit.remove-duplicates", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B3",
            keyColumns,
            hasHeaders = false
        });
        using var result = JsonDocument.Parse(response.Result!);
        var remaining = _commands.GetValues(_fixture.BatchToken, sheet, "B1:B3");
        Assert.True(remaining.Success, remaining.ErrorMessage);
        Assert.Equal(remaining.Values.Count(row => row[0] is not null),
            result.RootElement.GetProperty("remainingRows").GetInt32());
        Assert.Equal("first", remaining.Values[0][0]);
    }

    [Theory]
    [InlineData("""{"comma":true,"misspelled":true}""")]
    [InlineData("""{"comma":true,"qualifier":"unknown"}""")]
    [InlineData("""{"mode":"FixedWidth","comma":true,"fields":[{"position":0}]}""")]
    [InlineData("""{"mode":"FixedWidth","fields":[{"position":1}]}""")]
    [InlineData("""{"mode":"FixedWidth","fields":[{"position":0,"dataType":"Skip"}]}""")]
    [InlineData("""{"comma":true,"fields":[{"position":0}]}""")]
    [InlineData("""{"comma":true,"fields":[{"position":1},{"position":1}]}""")]
    [InlineData("""{"comma":true,"otherDelimiter":"too long"}""")]
    [InlineData("""{"comma":true,"decimalSeparator":".","thousandsSeparator":"."}""")]
    public async Task TextToColumns_InvalidOptionsDoNotWrite(string json)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1", [["x,y"]]).Success);
        var before = ReadNativeView();
        using var options = JsonDocument.Parse(json);
        var response = await _fixture.SendForFailureAsync("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1",
            destinationCell = "A1",
            options = options.RootElement
        });
        Assert.False(response.Success);
        Assert.Equal(before, ReadNativeView());
        var read = _commands.GetValues(_fixture.BatchToken, sheet, "A1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal("x,y", read.Values[0][0]);
    }

    [Fact]
    public async Task TextToColumns_EmptySourceNativeFailureStillClosesScratch()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var before = ReadNativeView();
        var response = await _fixture.SendForFailureAsync("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1:A3",
            destinationCell = "D1",
            options = new { comma = true }
        });
        Assert.False(response.Success);
        Assert.Equal(before, ReadNativeView());
    }

    [Fact]
    public async Task TextToColumns_PreflightsFormulaTextRatherThanItsShorterCalculatedValue()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheet, "A1",
            [["=SUBSTITUTE(\"x,y\",\",\",\"\")"]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "F1", [["protected"]]).Success);
        var response = await _fixture.SendForFailureAsync("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1",
            destinationCell = "D1",
            options = new { comma = true }
        });
        Assert.False(response.Success);
        Assert.Equal("Conflict", response.ErrorCategory);
        Assert.Contains("$F$1", response.ErrorMessage);
        var read = _commands.GetValues(_fixture.BatchToken, sheet, "D1:F1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Null(read.Values[0][0]);
        Assert.Equal("protected", read.Values[0][2]);
    }

    [Theory]
    [InlineData("quoted")]
    [InlineData("numeric-text")]
    [InlineData("empty-formula")]
    public void RemoveDuplicates_PreviewPreservesNativeKeyTypes(string kind)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        List<List<object?>> values = kind switch
        {
            "quoted" => [["''abc", "first"], ["'abc", "second"], ["''abc", "third"]],
            "numeric-text" => [["'01", "first"], [1, "second"], ["'01", "third"]],
            "empty-formula" => [[null, "first"], [null, "second"], [null, "third"]],
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B3", values).Success);
        if (kind == "empty-formula")
            Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheet, "A2:A3", [["=\"\""], ["=\"\""]]).Success);
        int[] keyColumns = [1];
        var response = _fixture.Send("rangeedit.remove-duplicates", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B3",
            keyColumns,
            hasHeaders = false
        });
        using var result = JsonDocument.Parse(response.Result!);
        var remaining = _commands.GetValues(_fixture.BatchToken, sheet, "B1:B3");
        Assert.True(remaining.Success, remaining.ErrorMessage);
        Assert.Equal(remaining.Values.Count(row => row[0] is not null),
            result.RootElement.GetProperty("remainingRows").GetInt32());
        Assert.Equal("first", remaining.Values[0][0]);
    }

    [Theory]
    [InlineData("")]
    [InlineData("0")]
    [InlineData("3")]
    [InlineData("1,1")]
    public async Task RemoveDuplicates_InvalidKeysDoNotWrite(string columns)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B2",
            [[1, "first"], [1, "second"]]).Success);
        int[] keyColumns = columns.Length == 0 ? [] : columns.Split(',').Select(int.Parse).ToArray();
        var before = ReadNativeView();
        var response = await _fixture.SendForFailureAsync("rangeedit.remove-duplicates", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B2",
            keyColumns,
            hasHeaders = false
        });
        Assert.False(response.Success);
        Assert.Equal(before, ReadNativeView());
        var read = _commands.GetValues(_fixture.BatchToken, sheet, "B1:B2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal("first", read.Values[0][0]);
        Assert.Equal("second", read.Values[1][0]);
    }

    [Fact]
    public async Task TextToColumns_FullWorksheetWidthCannotSilentlyTruncateAdditionalFields()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        string input = "x" + new string(',', 16_384);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1", [[input]]).Success);
        var before = ReadNativeView();
        var response = await _fixture.SendForFailureAsync("rangeedit.text-to-columns", new
        {
            sheetName = sheet,
            sourceRange = "A1",
            destinationCell = "A1",
            options = new { comma = true }
        });
        Assert.False(response.Success);
        Assert.Equal(before, ReadNativeView());
        var read = _commands.GetValues(_fixture.BatchToken, sheet, "A1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(input, read.Values[0][0]);
    }

    private (int Workbooks, string Window, string Sheet, string Selection) ReadNativeView()
    {
        (int, string, string, string) result = default;
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Workbooks? books = null;
            Excel.Window? window = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? selection = null;
            try
            {
                books = context.App.Workbooks;
                window = context.App.ActiveWindow;
                sheet = (Excel.Worksheet)context.App.ActiveSheet;
                selection = (Excel.Range)context.App.Selection;
                result = (books.Count, window.Caption, sheet.Name, selection.Address);
            }
            finally
            {
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref books);
            }
        });
        return result;
    }
}
