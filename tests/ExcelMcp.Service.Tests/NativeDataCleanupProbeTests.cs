using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class NativeDataCleanupProbeTests(
    PersistentServiceWorkbookFixture fixture, ITestOutputHelper output) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Theory]
    [InlineData("x,,")]
    [InlineData("\"x,y\",z")]
    [InlineData("x,y,z")]
    public void TextToColumns_ReturnTypeAndEmptyFieldsAreObservedNatively(string input)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[input]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "D1:F1",
            [["sentinel", "sentinel", "sentinel"]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? source = null;
            Excel.Range? destination = null;
            object? outcome = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                source = sheet.Range["A1"];
                destination = sheet.Range["D1"];
                outcome = source.TextToColumns(
                    Destination: destination,
                    DataType: Excel.XlTextParsingType.xlDelimited,
                    TextQualifier: Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
                    ConsecutiveDelimiter: false,
                    Tab: false, Semicolon: false, Comma: true, Space: false, Other: false,
                    DecimalSeparator: ".", ThousandsSeparator: ",", TrailingMinusNumbers: true);
                Assert.IsType<bool>(outcome);
                output.WriteLine($"input={input}; returned type={outcome?.GetType().FullName ?? "null"}; " +
                    $"returned range={(outcome as Excel.Range)?.Address ?? "not a range"}");
            }
            finally
            {
                ComUtilities.Release(ref outcome);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref sheet);
            }
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "D1:F1");
        Assert.True(read.Success, read.ErrorMessage);
        object?[] expected = input switch
        {
            "x,," => ["x", null, null],
            "\"x,y\",z" => ["x,y", "z", "sentinel"],
            _ => ["x", "y", "z"]
        };
        Assert.Equal(expected, Assert.Single(read.Values));
        var sourceRead = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
        Assert.True(sourceRead.Success, sourceRead.ErrorMessage);
        Assert.Equal(input, Assert.Single(Assert.Single(sourceRead.Values)));
        output.WriteLine(string.Join(" | ", read.Values[0].Select(value =>
            Convert.ToString(value, CultureInfo.InvariantCulture) ?? "<null>")));
    }

    [Fact]
    public void RemoveDuplicates_RemainingRowsAndCellsBelowTheRangeAreObservedNatively()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:B6",
            [["Key", "Amount"], [1, 10], [1, 20], [2, 30], [2, 30], ["below", "keep"]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                range = sheet.Range["A1:B5"];
                object[] columns = [1];
                range.RemoveDuplicates(columns, Excel.XlYesNoGuess.xlYes);
                output.WriteLine($"range after native removal={range.Address}");
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:B6");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(6, read.Values.Count);
        object?[][] expected = [["Key", "Amount"], [1, 10], [2, 30], [null, null], [null, null], ["below", "keep"]];
        for (var row = 0; row < expected.Length; row++)
            Assert.Equal(expected[row], read.Values[row]);
        foreach (var row in read.Values)
            output.WriteLine(string.Join(" | ", row.Select(value =>
                Convert.ToString(value, CultureInfo.InvariantCulture) ?? "<null>")));
    }

    [Fact]
    public void TextToColumns_WidestRowClearsFirstRowAcrossTheCompleteOutput()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A2",
            [["x"], ["x,y,z"]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "D1:G2",
            [["marker", "marker", "marker", "marker"], ["marker", "marker", "marker", "marker"]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? source = null;
            Excel.Range? destination = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                source = sheet.Range["A1:A2"];
                destination = sheet.Range["D1"];
                source.TextToColumns(Destination: destination, DataType: Excel.XlTextParsingType.xlDelimited,
                    TextQualifier: Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
                    ConsecutiveDelimiter: false, Tab: false, Semicolon: false, Comma: true,
                    Space: false, Other: false, DecimalSeparator: ".", ThousandsSeparator: ",");
            }
            finally
            {
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref sheet);
            }
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "D1:G2");
        Assert.True(read.Success, read.ErrorMessage);
        foreach (var row in read.Values)
            output.WriteLine(string.Join(" | ", row.Select(value =>
                Convert.ToString(value, CultureInfo.InvariantCulture) ?? "<null>")));
        Assert.Equal(2, read.Values.Count);
        Assert.Equal(["x", null, null, "marker"], read.Values[0]);
        Assert.Equal(["x", "y", "z", "marker"], read.Values[1]);
    }
}
