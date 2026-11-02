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
public sealed class PersistentServiceRangeErrorIdentityTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Theory]
    [InlineData(-2146826281, "get-values", "a1")]
    [InlineData(-2146826243, "get-values", "a1")]
    [InlineData(-2146826281, "get-formulas", "a1")]
    [InlineData(-2146826243, "get-formulas", "a1")]
    [InlineData(-2146826281, "get-formulas", "r1c1")]
    [InlineData(-2146826243, "get-formulas", "r1c1")]
    public void Reads_PreserveNumericErrorCodeCollisionsAndReportOnlyNativeErrors(
        int number, string action, string referenceStyle)
    {
        var sheetName = CreateErrorIdentitySheet(number);
        AssertNativeContents(sheetName, number);

        foreach (string address in new[] { "B3", "C3", "D3", "B3:D3" })
        {
            using var result = Read(action, sheetName, address, referenceStyle);
            var root = result.RootElement;
            bool multi = address == "B3:D3";
            bool hasError = multi || address == "D3";
            Assert.Equal(1, root.GetProperty("rowCount").GetInt32());
            Assert.Equal(multi ? 3 : 1, root.GetProperty("columnCount").GetInt32());
            var values = Assert.Single(root.GetProperty("values").EnumerateArray());
            Assert.Equal(multi ? 3 : 1, values.GetArrayLength());
            if (multi)
            {
                AssertNumeric(values[0], number);
                AssertNumeric(values[1], number);
            }
            else if (!hasError)
                AssertNumeric(values[0], number);

            var errors = root.GetProperty("cellErrors");
            if (hasError)
            {
                Assert.Equal("#DIV/0!", values[multi ? 2 : 0].GetString());
                var error = Assert.Single(errors.EnumerateArray());
                AssertDivisionByZeroError(error);
            }
            else
                Assert.Empty(errors.EnumerateArray());

            if (action == "get-formulas")
            {
                var formulas = Assert.Single(root.GetProperty("formulas").EnumerateArray());
                string reference = referenceStyle == "r1c1" ? "=RC[-1]" : "=B3";
                string[] expected = multi ? ["", reference, "=1/0"]
                    : [address == "B3" ? "" : address == "C3" ? reference : "=1/0"];
                Assert.Equal(expected, formulas.EnumerateArray().Select(value => value.GetString()));
            }
        }

        AssertNativeContents(sheetName, number);
    }

    [Theory]
    [InlineData("get-values", "a1")]
    [InlineData("get-formulas", "a1")]
    [InlineData("get-formulas", "r1c1")]
    public void Reads_MixedBlockChecksEachCellAndReportsAbsoluteErrorPositions(
        string action, string referenceStyle)
    {
        const int number = -2146826281;
        const int spillNumber = -2146826243;
        var sheetName = CreateErrorIdentitySheet(number);
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B4", [["=NA()"]]));
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "C4", [[spillNumber]]));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "D4", [["=C4"]]));
        AssertNativeContents(sheetName, number, hasSecondRow: true);

        using var result = Read(action, sheetName, "B3:D4", referenceStyle);
        var root = result.RootElement;
        Assert.Equal(2, root.GetProperty("rowCount").GetInt32());
        Assert.Equal(3, root.GetProperty("columnCount").GetInt32());
        var values = root.GetProperty("values");
        Assert.Equal(2, values.GetArrayLength());
        Assert.All(values.EnumerateArray(), row => Assert.Equal(3, row.GetArrayLength()));
        AssertNumeric(values[0][0], number);
        AssertNumeric(values[0][1], number);
        Assert.Equal("#DIV/0!", values[0][2].GetString());
        Assert.Equal("#N/A", values[1][0].GetString());
        AssertNumeric(values[1][1], spillNumber);
        AssertNumeric(values[1][2], spillNumber);
        Assert.Collection(root.GetProperty("cellErrors").EnumerateArray(),
            AssertDivisionByZeroError,
            error =>
            {
                Assert.Equal("B4", error.GetProperty("cellAddress").GetString());
                Assert.Equal(4, error.GetProperty("row").GetInt32());
                Assert.Equal(2, error.GetProperty("column").GetInt32());
                Assert.Equal("#N/A", error.GetProperty("errorName").GetString());
                Assert.Equal("=NA()", error.GetProperty("formula").GetString());
                Assert.Equal(-2146826246, error.GetProperty("currentValue").GetInt32());
                Assert.Equal(-2146826246, error.GetProperty("errorCode").GetInt32());
                Assert.Equal("#N/A - Value not available", error.GetProperty("errorMessage").GetString());
                Assert.Equal("Check that the lookup value and source data are available.",
                    error.GetProperty("suggestion").GetString());
            });
        if (action == "get-formulas")
        {
            var formulas = root.GetProperty("formulas");
            Assert.Equal(2, formulas.GetArrayLength());
            Assert.Equal(new[] { "", referenceStyle == "r1c1" ? "=RC[-1]" : "=B3", "=1/0" },
                formulas[0].EnumerateArray().Select(value => value.GetString()));
            Assert.Equal(new[] { "=NA()", "", referenceStyle == "r1c1" ? "=RC[-1]" : "=C4" },
                formulas[1].EnumerateArray().Select(value => value.GetString()));
        }

        AssertNativeContents(sheetName, number, hasSecondRow: true);
    }

    [Theory]
    [InlineData(-2146826281, "trace-precedents")]
    [InlineData(-2146826243, "trace-precedents")]
    [InlineData(-2146826281, "trace-dependents")]
    [InlineData(-2146826243, "trace-dependents")]
    public void TracePreviews_PreserveNumericErrorCodeCollisions(int number, string action)
    {
        var sheetName = CreateErrorIdentitySheet(number);
        using var result = Read(action, sheetName, "B3:D3");
        var nodes = result.RootElement.GetProperty("nodes").EnumerateArray().ToArray();
        Assert.Equal(3, nodes.Length);
        foreach (string address in new[] { "$B$3", "$C$3" })
            AssertNumeric(Assert.Single(nodes, node => node.GetProperty("address").GetString() == address)
                .GetProperty("value"), number);
        Assert.Equal("#DIV/0!", Assert.Single(nodes,
            node => node.GetProperty("address").GetString() == "$D$3").GetProperty("value").GetString());
        var edge = Assert.Single(result.RootElement.GetProperty("edges").EnumerateArray());
        Assert.Equal(action == "trace-precedents" ? "$C$3" : "$B$3",
            edge.GetProperty("fromAddress").GetString());
        Assert.Equal(action == "trace-precedents" ? "$B$3" : "$C$3",
            edge.GetProperty("toAddress").GetString());
        AssertNativeContents(sheetName, number);
    }

    [Theory]
    [InlineData(-2146826281, "trace-precedents")]
    [InlineData(-2146826243, "trace-precedents")]
    [InlineData(-2146826281, "trace-dependents")]
    [InlineData(-2146826243, "trace-dependents")]
    public void TracePreviews_DiscoveredCellsPreserveNumericErrorCodeCollisions(int number, string action)
    {
        var sheetName = CreateErrorIdentitySheet(number);
        bool precedents = action == "trace-precedents";
        string rootAddress = precedents ? "$C$3" : "$B$3";
        string discoveredAddress = precedents ? "$B$3" : "$C$3";
        using var result = Read(action, sheetName, precedents ? "C3" : "B3");
        var nodes = result.RootElement.GetProperty("nodes").EnumerateArray().ToArray();
        Assert.Equal(2, nodes.Length);
        var root = Assert.Single(nodes, node => node.GetProperty("address").GetString() == rootAddress);
        Assert.True(root.GetProperty("isRoot").GetBoolean());
        Assert.Equal(0, root.GetProperty("depth").GetInt32());
        AssertNumeric(root.GetProperty("value"), number);
        var discovered = Assert.Single(nodes, node => node.GetProperty("address").GetString() == discoveredAddress);
        Assert.False(discovered.GetProperty("isRoot").GetBoolean());
        Assert.Equal(1, discovered.GetProperty("depth").GetInt32());
        AssertNumeric(discovered.GetProperty("value"), number);
        var edge = Assert.Single(result.RootElement.GetProperty("edges").EnumerateArray());
        Assert.Equal(rootAddress, edge.GetProperty("fromAddress").GetString());
        Assert.Equal(discoveredAddress, edge.GetProperty("toAddress").GetString());
        AssertNativeContents(sheetName, number);
    }

    [Fact]
    public void SpillInspection_NumericSpillCodeFormulaIsOrdinaryAndBlockedSpillRemainsBlocked()
    {
        const int number = -2146826243;
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "D4", [["blocking"]]));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B3", [["=-2146826243"]]));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "D3", [["=SEQUENCE(3)"]]));

        using var result = Read("get-spill-info", sheetName, "B3,D3");
        var cells = result.RootElement.GetProperty("cells");
        Assert.Equal(2, cells.GetArrayLength());
        Assert.Equal("$B$3", cells[0].GetProperty("address").GetString());
        Assert.Equal("ordinary", cells[0].GetProperty("state").GetString());
        Assert.False(cells[0].TryGetProperty("sourceAddress", out _));
        Assert.Equal("$D$3", cells[1].GetProperty("address").GetString());
        Assert.Equal("blocked", cells[1].GetProperty("state").GetString());
        Assert.Equal("$D$3", cells[1].GetProperty("sourceAddress").GetString());
        Assert.Equal("=SEQUENCE(3)", cells[1].GetProperty("sourceFormula").GetString());
        Assert.False(cells[1].TryGetProperty("spillAddress", out _));
        var numeric = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheetName, "B3"));
        Assert.Equal("=-2146826243", numeric.Formulas[0][0]);
        Assert.Equal((double)number, Convert.ToDouble(numeric.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Empty(numeric.CellErrors);
        var blocked = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheetName, "D3"));
        Assert.Equal("=SEQUENCE(3)", blocked.Formulas[0][0]);
        Assert.Equal("#SPILL!", blocked.Values[0][0]);
        Assert.Equal("#SPILL!", Assert.Single(blocked.CellErrors).ErrorName);
        Assert.Equal("blocking", RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "D4")).Values[0][0]);
    }

    private string CreateErrorIdentitySheet(int number)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "B3", [[number]]));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "C3:D3", [["=B3", "=1/0"]]));
        return sheetName;
    }

    private JsonDocument Read(string action, string sheetName, string rangeAddress, string referenceStyle = "a1")
    {
        var response = action == "get-formulas"
            ? _fixture.Send($"range.{action}", new { sheetName, rangeAddress, referenceStyle })
            : _fixture.Send($"range.{action}", new { sheetName, rangeAddress });
        var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        return result;
    }

    private static void AssertNumeric(JsonElement value, int number)
    {
        Assert.Equal(JsonValueKind.Number, value.ValueKind);
        Assert.Equal((double)number, value.GetDouble());
    }

    private static void AssertDivisionByZeroError(JsonElement error)
    {
        Assert.Equal("D3", error.GetProperty("cellAddress").GetString());
        Assert.Equal(3, error.GetProperty("row").GetInt32());
        Assert.Equal(4, error.GetProperty("column").GetInt32());
        Assert.Equal("#DIV/0!", error.GetProperty("errorName").GetString());
        Assert.Equal("=1/0", error.GetProperty("formula").GetString());
        Assert.Equal(-2146826281, error.GetProperty("currentValue").GetInt32());
        Assert.Equal(-2146826281, error.GetProperty("errorCode").GetInt32());
        Assert.Equal("#DIV/0! - Division by zero", error.GetProperty("errorMessage").GetString());
        Assert.Equal("Ensure the formula does not divide by zero.", error.GetProperty("suggestion").GetString());
    }

    private void AssertNativeContents(string sheetName, int number, bool hasSecondRow = false)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                range = sheet.Range[hasSecondRow ? "B3:D4" : "B3:D3"];
                var values = Assert.IsType<object[,]>((object)range.Value2);
                Assert.Equal((double)number, Assert.IsType<double>(values[1, 1]));
                Assert.Equal((double)number, Assert.IsType<double>(values[1, 2]));
                Assert.Equal(-2146826281, Assert.IsType<int>(values[1, 3]));
                var formulas = Assert.IsType<object[,]>((object)range.Formula2);
                Assert.Equal("=B3", formulas[1, 2]);
                Assert.Equal("=1/0", formulas[1, 3]);
                if (hasSecondRow)
                {
                    Assert.Equal(-2146826246, Assert.IsType<int>(values[2, 1]));
                    Assert.Equal(-2146826243d, Assert.IsType<double>(values[2, 2]));
                    Assert.Equal(-2146826243d, Assert.IsType<double>(values[2, 3]));
                    Assert.Equal("=NA()", formulas[2, 1]);
                    Assert.Equal("=C4", formulas[2, 3]);
                }
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
