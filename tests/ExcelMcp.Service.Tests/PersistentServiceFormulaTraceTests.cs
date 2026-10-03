using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceFormulaTraceTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void Precedents_TraverseTheCompleteNativeLocalGraph()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[2]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1:D1",
            [["=A1*2", "=B1*2", "=SUM(B1,C1)"]]).Success);
        var response = _fixture.Send("range.trace-precedents", new { sheetName, rangeAddress = "D1" });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var nodes = result.RootElement.GetProperty("nodes").EnumerateArray().ToArray();
        Assert.Equal(4, nodes.Length);
        Assert.Equal(["$A$1", "$B$1", "$C$1", "$D$1"],
            nodes.Select(node => node.GetProperty("address").GetString()).Order(StringComparer.Ordinal));
        Assert.Equal(4, result.RootElement.GetProperty("edges").GetArrayLength());
        Assert.Equal("same-worksheet-only",
            result.RootElement.GetProperty("coverage").GetProperty("scope").GetString());
        Assert.Empty(result.RootElement.GetProperty("unresolved").EnumerateArray());
    }

    [Fact]
    public void Dependents_TraverseNativeDirectRelationshipsRatherThanFormulaText()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[2]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1:C1",
            [["=A1*2", "=B1*2"]]).Success);
        var response = _fixture.Send("range.trace-dependents", new { sheetName, rangeAddress = "A1" });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(3, result.RootElement.GetProperty("nodes").GetArrayLength());
        Assert.Equal(2, result.RootElement.GetProperty("edges").GetArrayLength());
        Assert.False(result.RootElement.GetProperty("coverage").GetProperty("workbookComplete").GetBoolean());
    }

    [Theory]
    [InlineData("=1+2")]
    [InlineData("=INDIRECT(\"A1\")")]
    public void NativeAbsence_IsExplicitlyUnresolvedRatherThanInventedEmpty(string formula)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[2]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1", [[formula]]).Success);
        var response = _fixture.Send("range.trace-precedents", new { sheetName, rangeAddress = "B1" });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Single(result.RootElement.GetProperty("nodes").EnumerateArray());
        Assert.Empty(result.RootElement.GetProperty("edges").EnumerateArray());
        var unresolved = Assert.Single(result.RootElement.GetProperty("unresolved").EnumerateArray());
        Assert.Equal("$B$1", unresolved.GetProperty("address").GetString());
        Assert.Equal("0x800A03EC", unresolved.GetProperty("nativeErrorCode").GetString());
        Assert.False(result.RootElement.GetProperty("coverage").GetProperty("nativeTraversalComplete").GetBoolean());
    }

    [Fact]
    public void MixedCrossSheetReferences_ReportOnlyNativeLocalEdgesAndLimitedCoverage()
    {
        var source = _fixture.CreateTestSheet(_fixture.BatchToken);
        var other = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, source, "A1", [[2]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, other, "A1", [[7]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, source, "B1", [[$"=A1+'{other}'!A1"]]).Success);
        var response = _fixture.Send("range.trace-precedents", new { sheetName = source, rangeAddress = "B1" });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal(2, result.RootElement.GetProperty("nodes").GetArrayLength());
        var edge = Assert.Single(result.RootElement.GetProperty("edges").EnumerateArray());
        Assert.Equal("$B$1", edge.GetProperty("fromAddress").GetString());
        Assert.Equal("$A$1", edge.GetProperty("toAddress").GetString());
        var coverage = result.RootElement.GetProperty("coverage");
        Assert.False(coverage.GetProperty("workbookComplete").GetBoolean());
        Assert.Contains(coverage.GetProperty("limitations").EnumerateArray(), limitation =>
            limitation.GetString()!.Contains("Cross-worksheet", StringComparison.Ordinal));
    }

    [Fact]
    public void ExactDisjointScopesAndNamedReferences_AreTraversedWithoutGapsOrDuplicates()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"TraceInput_{Guid.NewGuid():N}";
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[2]]).Success);
        _fixture.Send("namedrange.create", new { name, reference = $"'{sheetName}'!$A$1" });
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1:C1",
            [[$"={name}*2", "=B1*2"]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "D1", [["not requested"]]).Success);
        var response = _fixture.Send("range.trace-precedents", new
        {
            sheetName,
            rangeAddress = "B1:C1,C1"
        });
        using var result = JsonDocument.Parse(response.Result!);
        var nodes = result.RootElement.GetProperty("nodes").EnumerateArray().ToArray();
        Assert.Equal(3, nodes.Length);
        Assert.Equal(2, nodes.Count(node => node.GetProperty("isRoot").GetBoolean()));
        Assert.DoesNotContain(nodes, node => node.GetProperty("address").GetString() == "$D$1");
        Assert.Equal(2, result.RootElement.GetProperty("edges").GetArrayLength());
        var named = _fixture.Send("range.trace-precedents", new { sheetName = "", rangeAddress = name });
        using var namedResult = JsonDocument.Parse(named.Result!);
        Assert.Equal(sheetName, namedResult.RootElement.GetProperty("sheetName").GetString());
        Assert.Single(namedResult.RootElement.GetProperty("nodes").EnumerateArray());
        Assert.Empty(namedResult.RootElement.GetProperty("unresolved").EnumerateArray());
    }

    [Theory]
    [InlineData("trace-precedents")]
    [InlineData("trace-dependents")]
    public void CircularGraphs_TerminateAndIdentifyStronglyConnectedCells(string action)
    {
        var calculation = _fixture.CreateCommands<ICalculationModeCommands>();
        var previous = calculation.GetSettings(_fixture.BatchToken);
        Assert.True(previous.Success);
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        try
        {
            Assert.True(calculation.SetSettings(_fixture.BatchToken, CalculationMode.Manual).Success);
            Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1:B1", [["=B1", "=A1"]]).Success);
            var response = _fixture.Send($"range.{action}", new { sheetName, rangeAddress = "A1" });
            using var result = JsonDocument.Parse(response.Result!);
            Assert.Equal(2, result.RootElement.GetProperty("nodes").GetArrayLength());
            Assert.Equal(2, result.RootElement.GetProperty("edges").GetArrayLength());
            var cycle = Assert.Single(result.RootElement.GetProperty("cycles").EnumerateArray());
            Assert.Equal(["$A$1", "$B$1"], cycle.EnumerateArray().Select(cell => cell.GetString()));
        }
        finally
        {
            try
            {
                Assert.True(_commands.ClearContents(_fixture.BatchToken, sheetName, "A1:B1").Success);
            }
            finally
            {
                Assert.True(calculation.SetSettings(_fixture.BatchToken, (CalculationMode)previous.ModeValue).Success);
            }
        }
    }

    [Theory]
    [InlineData("trace-precedents", "B1")]
    [InlineData("trace-dependents", "A1")]
    public void TracingInactiveWorksheet_DoesNotChangeActiveSheetOrSelection(string action, string rangeAddress)
    {
        var source = _fixture.CreateTestSheet(_fixture.BatchToken);
        var active = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, source, "A1", [[2]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, source, "B1", [["=A1*2"]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? selected = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, active);
                Assert.NotNull(sheet);
                selected = sheet.Range["C5"];
                sheet.Activate();
                selected.Select();
            }
            finally
            {
                ComUtilities.Release(ref selected);
                ComUtilities.Release(ref sheet);
            }
        });
        var response = _fixture.Send($"range.{action}", new { sheetName = source, rangeAddress });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal(2, result.RootElement.GetProperty("nodes").GetArrayLength());
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? selection = null;
            try
            {
                sheet = context.Book.ActiveSheet as Excel.Worksheet;
                selection = context.App.Selection as Excel.Range;
                Assert.NotNull(sheet);
                Assert.NotNull(selection);
                Assert.Equal(active, sheet.Name);
                Assert.Equal("$C$5", selection.Address);
            }
            finally
            {
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    [Fact]
    public void StructuredTableReferences_UseNativeCellRelationships()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var tableName = $"TraceTable_{Guid.NewGuid():N}";
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A3",
            [["Amount"], [2], [3]]).Success);
        _fixture.Send("table.create", new { sheetName, tableName, rangeAddress = "A1:A3", hasHeaders = true });
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "C1", [[$"=SUM({tableName}[Amount])"]]).Success);
        var response = _fixture.Send("range.trace-precedents", new { sheetName, rangeAddress = "C1" });
        using var result = JsonDocument.Parse(response.Result!);
        var addresses = result.RootElement.GetProperty("nodes").EnumerateArray()
            .Select(node => node.GetProperty("address").GetString()).ToArray();
        Assert.Contains("$A$2", addresses);
        Assert.Contains("$A$3", addresses);
        Assert.Contains("$C$1", addresses);
        Assert.Empty(result.RootElement.GetProperty("unresolved").EnumerateArray());
    }
}
