using System.Globalization;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceRangeFillTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void Fill_DownAdjustsNativeRelativeReferences()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A3", [[1], [2], [3]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1", [["=A1*2"]]).Success);
        _fixture.Send("rangeedit.fill", new { sheetName, rangeAddress = "B1:B3", direction = "down" });
        var read = _commands.GetFormulas(_fixture.BatchToken, sheetName, "B1:B3");
        Assert.True(read.Success);
        Assert.Equal(["=A1*2", "=A2*2", "=A3*2"], read.Formulas.Select(row => row[0]));
        Assert.Equal([2d, 4d, 6d], read.Values.Select(row => Convert.ToDouble(row[0], CultureInfo.InvariantCulture)));
    }

    [Fact]
    public async Task Fill_RejectsOccupiedDestinationWithoutReplacingSourceOrOtherCells()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[1]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A3", [["keep"]]).Success);
        var failure = await _fixture.SendForFailureAsync("rangeedit.fill", new
        {
            sheetName,
            rangeAddress = "A1:A3",
            direction = "down"
        });
        Assert.Equal("Conflict", failure.ErrorCategory);
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:A3");
        Assert.True(read.Success);
        Assert.Equal(1, Convert.ToDouble(read.Values[0][0], CultureInfo.InvariantCulture));
        Assert.Null(read.Values[1][0]);
        Assert.Equal("keep", read.Values[2][0]);
    }

    [Fact]
    public void AutoFill_ExtendsNativeNumberPatternWithoutRequiringOverwriteOfSource()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A2", [[1], [3]]).Success);
        _fixture.Send("rangeedit.auto-fill", new
        {
            sheetName,
            sourceRange = "A1:A2",
            destinationRange = "A1:A5",
            fillType = "series"
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:A5");
        Assert.True(read.Success);
        Assert.Equal([1d, 3d, 5d, 7d, 9d], read.Values.Select(row => Convert.ToDouble(row[0], CultureInfo.InvariantCulture)));
    }

    [Fact]
    public void CreateSeries_UsesNativeLinearSteps()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[10]]).Success);
        _fixture.Send("rangeedit.create-series", new
        {
            sheetName,
            rangeAddress = "A1:A4",
            orientation = "columns",
            seriesType = "linear",
            stepValue = 5d
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:A4");
        Assert.True(read.Success);
        Assert.Equal([10d, 15d, 20d, 25d], read.Values.Select(row => Convert.ToDouble(row[0], CultureInfo.InvariantCulture)));
    }

    [Fact]
    public void Formulas_R1C1WriteAndReadUseNativeRelativeNotation()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A2", [[5], [7]]).Success);
        List<List<string>> formulas = [["=RC[-1]*2"], ["=RC[-1]*2"]];
        _fixture.Send("range.set-formulas", new
        {
            sheetName,
            rangeAddress = "B1:B2",
            formulas,
            referenceStyle = "r1c1"
        });
        var response = _fixture.Send("range.get-formulas", new
        {
            sheetName,
            rangeAddress = "B1:B2",
            referenceStyle = "r1c1"
        });
        using var read = JsonDocument.Parse(response.Result!);
        Assert.Equal("=RC[-1]*2", read.RootElement.GetProperty("formulas")[0][0].GetString());
        Assert.Equal("=RC[-1]*2", read.RootElement.GetProperty("formulas")[1][0].GetString());
        var a1 = _commands.GetFormulas(_fixture.BatchToken, sheetName, "B1:B2");
        Assert.True(a1.Success);
        Assert.Equal(["=A1*2", "=A2*2"], a1.Formulas.Select(row => row[0]));
        Assert.Equal([10d, 14d], a1.Values.Select(row => Convert.ToDouble(row[0], CultureInfo.InvariantCulture)));
    }

    [Theory]
    [InlineData("down", "A1", "A1:A3")]
    [InlineData("up", "A3", "A1:A3")]
    [InlineData("left", "C1", "A1:C1")]
    [InlineData("right", "A1", "A1:C1")]
    public void Fill_AllDirectionsCopyNativeContentAndFormats(
        string direction, string seed, string destination)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, seed, [[42]]).Success);
        _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])[seed],
            formatOptions = new
            {
                bold = true
            }
        });
        _fixture.Send("rangeedit.fill", new { sheetName, rangeAddress = destination, direction });
        var values = _commands.GetValues(_fixture.BatchToken, sheetName, destination);
        Assert.True(values.Success);
        Assert.All(values.Values.SelectMany(row => row), value =>
            Assert.Equal(42d, Convert.ToDouble(value, CultureInfo.InvariantCulture)));
        var response = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress = destination });
        using var format = JsonDocument.Parse(response.Result!);
        Assert.Equal(3, format.RootElement.GetProperty("cellCount").GetInt64());
        Assert.All(format.RootElement.GetProperty("cells").EnumerateArray(), cell =>
            Assert.True(cell.GetProperty("stored").GetProperty("font").GetProperty("bold").GetBoolean()));
    }

    [Theory]
    [InlineData("A1:A2", "A1:A5", false, 1)]
    [InlineData("A4:A5", "A1:A5", false, 7)]
    [InlineData("A1:B1", "A1:E1", true, 1)]
    [InlineData("D1:E1", "A1:E1", true, 7)]
    public void AutoFill_ExtendsInAllFourDirections(string sourceRange, string destinationRange,
        bool horizontal, int start)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        List<List<object?>> seeds = horizontal ? [[start, start + 2]] : [[start], [start + 2]];
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, sourceRange, seeds).Success);
        _fixture.Send("rangeedit.auto-fill", new
        {
            sheetName,
            sourceRange,
            destinationRange,
            fillType = "series"
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, destinationRange);
        Assert.True(read.Success);
        Assert.Equal([1d, 3d, 5d, 7d, 9d],
            read.Values.SelectMany(row => row).Select(value => Convert.ToDouble(value, CultureInfo.InvariantCulture)));
    }

    [Theory]
    [InlineData("A1:A2", "B1:B5")]
    [InlineData("A1:A2", "A1:B5")]
    [InlineData("A2:A3", "A1:A5")]
    [InlineData("A1:A2", "A1:A2")]
    public async Task AutoFill_RejectsAmbiguousDestinationBeforeWriting(string sourceRange, string destinationRange)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, sourceRange, [[1], [3]]).Success);
        var before = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:B5");
        RequireSuccess(before);
        var failure = await _fixture.SendForFailureAsync("rangeedit.auto-fill", new
        {
            sheetName,
            sourceRange,
            destinationRange,
            overwritePolicy = "allow"
        });
        Assert.Equal("InvalidInput", failure.ErrorCategory);
        var source = _commands.GetValues(_fixture.BatchToken, sheetName, sourceRange);
        Assert.True(source.Success);
        Assert.Equal([1d, 3d], source.Values.Select(row => Convert.ToDouble(row[0], CultureInfo.InvariantCulture)));
        var after = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:B5");
        RequireSuccess(after);
        for (int row = 0; row < before.Values.Count; row++)
            Assert.Equal(before.Values[row], after.Values[row]);
    }

    [Fact]
    public void AutoFill_FormatsPreservesOccupiedContent()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A3", [[1], ["keep"], [3]]).Success);
        _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1"],
            formatOptions = new
            {
                bold = true
            }
        });
        _fixture.Send("rangeedit.auto-fill", new
        {
            sheetName,
            sourceRange = "A1",
            destinationRange = "A1:A3",
            fillType = "formats"
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:A3");
        Assert.True(read.Success);
        Assert.Equal(1d, Convert.ToDouble(read.Values[0][0], CultureInfo.InvariantCulture));
        Assert.Equal("keep", read.Values[1][0]);
        Assert.Equal(3d, Convert.ToDouble(read.Values[2][0], CultureInfo.InvariantCulture));
        var response = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress = "A1:A3" });
        using var format = JsonDocument.Parse(response.Result!);
        Assert.Equal(3, format.RootElement.GetProperty("cells").GetArrayLength());
        Assert.All(format.RootElement.GetProperty("cells").EnumerateArray(), cell =>
            Assert.True(cell.GetProperty("stored").GetProperty("font").GetProperty("bold").GetBoolean()));
    }

    [Theory]
    [InlineData("rows", "A1:D1", "linear", 2d, 3d, 5d, 7d, 9d)]
    [InlineData("columns", "A1:A4", "growth", 2d, 3d, 6d, 12d, 24d)]
    public void Series_UsesNativeOrientationAndProgression(string orientation, string rangeAddress,
        string seriesType, double stepValue, double seed, double second, double third, double fourth)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[seed]]).Success);
        _fixture.Send("rangeedit.create-series", new
        {
            sheetName,
            rangeAddress,
            orientation,
            seriesType,
            stepValue
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, rangeAddress);
        Assert.True(read.Success);
        Assert.Equal([seed, second, third, fourth],
            read.Values.SelectMany(row => row).Select(value => Convert.ToDouble(value, CultureInfo.InvariantCulture)));
    }

    [Fact]
    public void Series_DateMonthsAndStoppingValueUseNativeSemantics()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var seed = new DateTime(2025, 1, 31).ToOADate();
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[seed]]).Success);
        _fixture.Send("rangeedit.create-series", new
        {
            sheetName,
            rangeAddress = "A1:A3",
            orientation = "columns",
            seriesType = "date",
            dateUnit = "month",
            stepValue = 1d
        });
        var dates = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:A3");
        Assert.True(dates.Success);
        Assert.Equal([seed, new DateTime(2025, 2, 28).ToOADate(), new DateTime(2025, 3, 31).ToOADate()],
            dates.Values.Select(row => Convert.ToDouble(row[0], CultureInfo.InvariantCulture)));
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "C1", [[1]]).Success);
        _fixture.Send("rangeedit.create-series", new
        {
            sheetName,
            rangeAddress = "C1:C5",
            orientation = "columns",
            stepValue = 2d,
            stopValue = 5d
        });
        var series = _commands.GetValues(_fixture.BatchToken, sheetName, "C1:C5");
        Assert.True(series.Success);
        Assert.Equal([1d, 3d, 5d], series.Values.Take(3).Select(row => Convert.ToDouble(row[0], CultureInfo.InvariantCulture)));
        Assert.Null(series.Values[3][0]);
        Assert.Null(series.Values[4][0]);
    }

    [Theory]
    [InlineData("rangeedit.fill", "A1,A3")]
    [InlineData("rangeedit.create-series", "A1,A3")]
    public async Task FillOperations_RejectDisjointRangesEvenWithAllow(string command, string rangeAddress)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[1]]).Success);
        var failure = await _fixture.SendForFailureAsync(command, new
        {
            sheetName,
            rangeAddress,
            direction = "down",
            orientation = "columns",
            overwritePolicy = "allow"
        });
        Assert.Equal("InvalidInput", failure.ErrorCategory);
        var unchanged = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:A3");
        Assert.True(unchanged.Success);
        Assert.Equal(1d, Convert.ToDouble(unchanged.Values[0][0], CultureInfo.InvariantCulture));
        Assert.Null(unchanged.Values[1][0]);
        Assert.Null(unchanged.Values[2][0]);
    }
}
