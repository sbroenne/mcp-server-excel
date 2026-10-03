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
public sealed class PersistentServiceRangeSpillTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void GetSpillInfo_UsesNativeSourceAndResultRelationships()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B2",
            [["=SEQUENCE(3,2)"]]).Success);
        AssertNativeSpill(sheetName, "C3", "$B$2", "$B$2:$C$4");
        using var result = ReadSpill(sheetName, "B2:C4");
        Assert.Equal("supported", result.RootElement.GetProperty("capability").GetString());
        var cells = result.RootElement.GetProperty("cells");
        Assert.Equal(6, cells.GetArrayLength());
        Assert.Equal("source", cells[0].GetProperty("state").GetString());
        Assert.Equal("result", cells[1].GetProperty("state").GetString());
        foreach (var cell in cells.EnumerateArray())
        {
            Assert.Equal("$B$2", cell.GetProperty("sourceAddress").GetString());
            Assert.Equal("$B$2:$C$4", cell.GetProperty("spillAddress").GetString());
            Assert.Equal("=SEQUENCE(3,2)", cell.GetProperty("sourceFormula").GetString());
            Assert.Equal(3, cell.GetProperty("spillRows").GetInt32());
            Assert.Equal(2, cell.GetProperty("spillColumns").GetInt32());
        }
        using var subset = ReadSpill(sheetName, "C3");
        Assert.Single(subset.RootElement.GetProperty("cells").EnumerateArray());
        Assert.Equal("$B$2", subset.RootElement.GetProperty("cells")[0]
            .GetProperty("sourceAddress").GetString());
    }

    [Fact]
    public void GetSpillInfo_BlockedFormulaDoesNotInventAnExtent()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A2", [["blocking"]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1",
            [["=SEQUENCE(3)"]]).Success);
        var values = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
        Assert.True(values.Success);
        Assert.Equal("#SPILL!", values.Values[0][0]);
        using var result = ReadSpill(sheetName, "A1:A3");
        var cells = result.RootElement.GetProperty("cells");
        Assert.Equal("blocked", cells[0].GetProperty("state").GetString());
        Assert.Equal("$A$1", cells[0].GetProperty("sourceAddress").GetString());
        Assert.False(cells[0].TryGetProperty("spillAddress", out _));
        Assert.False(cells[0].TryGetProperty("spillRows", out _));
        Assert.Equal("ordinary", cells[1].GetProperty("state").GetString());
        Assert.Equal("ordinary", cells[2].GetProperty("state").GetString());
    }

    [Fact]
    public void GetSpillInfo_FollowsChangingSizeAndIncludesEveryRequestedOrdinaryCell()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[3]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1",
            [["=SEQUENCE(A1)"]]).Success);
        using var first = ReadSpill(sheetName, "B1");
        Assert.Equal("$B$1:$B$3", first.RootElement.GetProperty("cells")[0].GetProperty("spillAddress").GetString());
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[5]],
            overwritePolicy: Sbroenne.ExcelMcp.Core.Commands.Range.OverwritePolicy.Allow).Success);
        using var changed = ReadSpill(sheetName, "B1");
        Assert.Equal("$B$1:$B$5", changed.RootElement.GetProperty("cells")[0].GetProperty("spillAddress").GetString());
        using var complete = ReadSpill(sheetName, "D1:D40");
        Assert.Equal(40, complete.RootElement.GetProperty("cellCount").GetInt64());
        Assert.All(complete.RootElement.GetProperty("cells").EnumerateArray(),
            cell => Assert.Equal("ordinary", cell.GetProperty("state").GetString()));
    }

    private JsonDocument ReadSpill(string sheetName, string rangeAddress)
    {
        var response = _fixture.Send("range.get-spill-info", new { sheetName, rangeAddress });
        var document = JsonDocument.Parse(response.Result!);
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        return document;
    }

    [Fact]
    public void GetSpillInfo_ResolvesNamesAndDisjointScopesWithoutChangingSelection()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1",
            [["=SEQUENCE(3)"]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "C1",
            [["=1+1"]]).Success);
        var name = $"Spill_{Guid.NewGuid():N}";
        _fixture.Send("namedrange.create", new { name, reference = $"'{sheetName}'!$A$2:$A$3" });
        _fixture.RegisterNamedRangeForCleanup(name);
        var before = ReadViewState();
        using var named = ReadSpill("", name);
        Assert.Equal(sheetName, named.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal(2, named.RootElement.GetProperty("cells").GetArrayLength());
        Assert.All(named.RootElement.GetProperty("cells").EnumerateArray(),
            cell => Assert.Equal("result", cell.GetProperty("state").GetString()));
        using var union = ReadSpill(sheetName, "A1:A2,A2:A3,C1");
        var cells = union.RootElement.GetProperty("cells");
        Assert.Equal(4, cells.GetArrayLength());
        Assert.Equal(["$A$1", "$C$1", "$A$2", "$A$3"],
            cells.EnumerateArray().Select(cell => cell.GetProperty("address").GetString()));
        Assert.Equal("ordinary", cells[1].GetProperty("state").GetString());
        Assert.False(cells[1].TryGetProperty("sourceAddress", out _));
        Assert.Equal(before, ReadViewState());
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

    private void AssertNativeSpill(string sheetName, string address, string sourceAddress, string spillAddress)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            Excel.Range? source = null;
            Excel.Range? spill = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range[address];
                Assert.True(Convert.ToBoolean((object)cell.HasSpill, System.Globalization.CultureInfo.InvariantCulture));
                source = cell.SpillParent;
                spill = source.SpillingToRange;
                Assert.Equal(sourceAddress, source.Address);
                Assert.Equal(spillAddress, spill.Address);
            }
            finally
            {
                ComUtilities.Release(ref spill);
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
