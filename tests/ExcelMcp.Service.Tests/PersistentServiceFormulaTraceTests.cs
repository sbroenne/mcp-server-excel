using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Commands.Range;
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
    private static readonly int[] PreservationDepths = [3, 0];

    [Theory]
    [InlineData("trace-precedents")]
    [InlineData("trace-dependents")]
    public void Limits_OmittedArgumentsUseShortReportDefaults(string action)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string text = new('x', 257);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[text]]));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1:E1",
            [["=A1", "=B1", "=C1", "=D1"]]));
        var response = _fixture.Send($"range.{action}", new
        {
            sheetName,
            rangeAddress = action == "trace-precedents" ? "E1" : "A1"
        });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal(4, result.RootElement.GetProperty("nodes").GetArrayLength());
        Assert.True(result.RootElement.GetProperty("truncation").GetProperty("depthLimitReached").GetBoolean());
        Assert.True(result.RootElement.GetProperty("truncation").GetProperty("previewTextTruncated").GetBoolean());
        Assert.All(result.RootElement.GetProperty("nodes").EnumerateArray(), node =>
        {
            Assert.Equal(new string('x', 256), node.GetProperty("value").GetString());
            Assert.True(node.GetProperty("valueTruncated").GetBoolean());
        });
        var rootsResponse = _fixture.Send($"range.{action}", new { sheetName, rangeAddress = "A1:A51" });
        using var roots = JsonDocument.Parse(rootsResponse.Result!);
        Assert.Equal(50, roots.RootElement.GetProperty("nodes").GetArrayLength());
        Assert.True(roots.RootElement.GetProperty("truncation").GetProperty("startingCellsOmitted").GetBoolean());
        Assert.True(roots.RootElement.GetProperty("truncation").GetProperty("nodeLimitReached").GetBoolean());
    }

    [Fact]
    public void Offset_ReportsTheNativeAnchorRatherThanInventingAnEffectiveTarget()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A2", [[2], [7]]));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1", [["=OFFSET(A1,1,0)"]]));
        var response = _fixture.Send("range.trace-precedents", new { sheetName, rangeAddress = "B1" });
        using var result = JsonDocument.Parse(response.Result!);
        AssertGraph(result, ["$B$1"], [("$B$1", "$A$1")]);
        Assert.Equal(7, result.RootElement.GetProperty("nodes")[1].GetProperty("value").GetInt32());
        Assert.False(result.RootElement.GetProperty("coverage").GetProperty("workbookComplete").GetBoolean());
        Assert.False(result.RootElement.GetProperty("truncation").GetProperty("any").GetBoolean());
    }

    [Fact]
    public void SpillReferences_KeepNativeCoverageAndSpilledValues()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1", [["=SEQUENCE(2)"]]));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1", [["=SUM(A1#)"]]));
        var response = _fixture.Send("range.trace-precedents", new { sheetName, rangeAddress = "B1", maxDepth = 1 });
        using var result = JsonDocument.Parse(response.Result!);
        AssertGraph(result, ["$B$1"], [("$B$1", "$A$1"), ("$B$1", "$A$2")]);
        Assert.Equal(3, result.RootElement.GetProperty("nodes").EnumerateArray()
            .Single(node => node.GetProperty("address").GetString() == "$B$1").GetProperty("value").GetInt32());
        Assert.Equal(1d, Convert.ToDouble(
            RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1")).Values[0][0],
            System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(2d, Convert.ToDouble(
            RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A2")).Values[0][0],
            System.Globalization.CultureInfo.InvariantCulture));
        Assert.False(result.RootElement.GetProperty("coverage").GetProperty("workbookComplete").GetBoolean());
    }

    [Fact]
    public void ClosedExternalReference_RemainsUnresolvedAndDoesNotOpenTheWorkbook()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string external = _fixture.CreateBlankWorkbook("trace", "closed-source");
        string formula = $"='{Path.GetDirectoryName(external)}\\[{Path.GetFileName(external)}]Sheet1'!$A$1";
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1", [[formula]]));
        int before = _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Workbooks? books = null;
            try
            {
                books = context.App.Workbooks;
                return books.Count;
            }
            finally
            {
                ComUtilities.Release(ref books);
            }
        });
        var response = _fixture.Send("range.trace-precedents", new { sheetName, rangeAddress = "B1", maxDepth = 1 });
        using var result = JsonDocument.Parse(response.Result!);
        AssertGraph(result, ["$B$1"], []);
        Assert.Single(result.RootElement.GetProperty("unresolved").EnumerateArray());
        Assert.False(result.RootElement.GetProperty("coverage").GetProperty("nativeTraversalComplete").GetBoolean());
        Assert.False(result.RootElement.GetProperty("coverage").GetProperty("workbookComplete").GetBoolean());
        Assert.False(result.RootElement.GetProperty("truncation").GetProperty("any").GetBoolean());
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Workbooks? books = null;
            try
            {
                books = context.App.Workbooks;
                Assert.Equal(before, books.Count);
                Assert.Equal(1, books.Count);
            }
            finally
            {
                ComUtilities.Release(ref books);
            }
        });
    }

    [Theory]
    [InlineData("trace-precedents", "D1", "$D$1", "$C$1")]
    [InlineData("trace-dependents", "A1", "$A$1", "$B$1")]
    public void Limits_StopAtTheRequestedDepth(string action, string rangeAddress, string root, string next)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[2]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1:D1",
            [["=A1*2", "=B1*2", "=C1*2"]]).Success);
        var response = _fixture.Send($"range.{action}", new
        {
            sheetName,
            rangeAddress,
            maxDepth = 1,
            maxNodes = 50,
            maxPreviewCharacters = 256
        });
        using var result = JsonDocument.Parse(response.Result!);
        AssertGraph(result, [root], [(root, next)]);
        Assert.Equal(0, result.RootElement.GetProperty("nodes").EnumerateArray()
            .Single(node => node.GetProperty("address").GetString() == root).GetProperty("depth").GetInt32());
        Assert.Equal(1, result.RootElement.GetProperty("nodes").EnumerateArray()
            .Single(node => node.GetProperty("address").GetString() == next).GetProperty("depth").GetInt32());
        var truncation = result.RootElement.GetProperty("truncation");
        Assert.True(truncation.GetProperty("depthLimitReached").GetBoolean());
        Assert.False(truncation.GetProperty("nodeLimitReached").GetBoolean());
        Assert.False(truncation.GetProperty("startingCellsOmitted").GetBoolean());
        Assert.False(truncation.GetProperty("previewTextTruncated").GetBoolean());
        Assert.True(truncation.GetProperty("any").GetBoolean());
        Assert.False(result.RootElement.GetProperty("coverage").GetProperty("nativeTraversalComplete").GetBoolean());
        var boundary = Assert.Single(result.RootElement.GetProperty("boundaries").EnumerateArray());
        Assert.Equal(next, boundary.GetProperty("address").GetString());
        Assert.Equal("max-depth", boundary.GetProperty("reason").GetString());
    }

    [Theory]
    [InlineData("trace-precedents")]
    [InlineData("trace-dependents")]
    public void Limits_OmitExcessStartingCellsWithoutWalkingTheWholeColumn(string action)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:A3", [[1], [2], [3]]).Success);
        var response = _fixture.Send($"range.{action}", new
        {
            sheetName,
            rangeAddress = "A:A",
            maxDepth = 0,
            maxNodes = 2
        });
        using var result = JsonDocument.Parse(response.Result!);
        AssertGraph(result, ["$A$1", "$A$2"], []);
        var truncation = result.RootElement.GetProperty("truncation");
        Assert.True(truncation.GetProperty("nodeLimitReached").GetBoolean());
        Assert.True(truncation.GetProperty("startingCellsOmitted").GetBoolean());
        Assert.True(truncation.GetProperty("any").GetBoolean());
        Assert.False(result.RootElement.GetProperty("coverage").GetProperty("nativeTraversalComplete").GetBoolean());
        Assert.Equal("$A:$A", result.RootElement.GetProperty("rangeAddress").GetString());
    }

    [Theory]
    [InlineData("trace-precedents", 0)]
    [InlineData("trace-precedents", 1)]
    [InlineData("trace-precedents", 3)]
    [InlineData("trace-precedents", 20)]
    [InlineData("trace-dependents", 0)]
    [InlineData("trace-dependents", 1)]
    [InlineData("trace-dependents", 3)]
    [InlineData("trace-dependents", 20)]
    public void Limits_DepthCountsStepsAndDoesNotFlagAnExactFit(string action, int maxDepth)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[2]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1:D1",
            [["=A1*2", "=B1*2", "=C1*2"]]).Success);
        bool precedents = action == "trace-precedents";
        var response = _fixture.Send($"range.{action}", new
        {
            sheetName,
            rangeAddress = precedents ? "D1" : "A1",
            maxDepth,
            maxNodes = 1000,
            maxPreviewCharacters = 4096
        });
        using var result = JsonDocument.Parse(response.Result!);
        int steps = Math.Min(maxDepth, 3);
        var nodes = result.RootElement.GetProperty("nodes").EnumerateArray().ToArray();
        Assert.Equal(steps + 1, nodes.Length);
        for (int depth = 0; depth <= steps; depth++)
        {
            string address = $"${(char)((precedents ? 'D' : 'A') + (precedents ? -depth : depth))}$1";
            Assert.Equal(depth, nodes.Single(node => node.GetProperty("address").GetString() == address)
                .GetProperty("depth").GetInt32());
        }
        Assert.Equal(steps, result.RootElement.GetProperty("edges").GetArrayLength());
        Assert.Equal(maxDepth < 3, result.RootElement.GetProperty("truncation").GetProperty("any").GetBoolean());
        Assert.Equal(maxDepth < 3, result.RootElement.GetProperty("truncation").GetProperty("depthLimitReached").GetBoolean());
        Assert.False(result.RootElement.GetProperty("truncation").GetProperty("nodeLimitReached").GetBoolean());
        var limits = result.RootElement.GetProperty("limits");
        Assert.Equal(maxDepth, limits.GetProperty("maxDepth").GetInt32());
        Assert.Equal(1000, limits.GetProperty("maxNodes").GetInt32());
        Assert.Equal(4096, limits.GetProperty("maxPreviewCharacters").GetInt32());
    }

    [Theory]
    [InlineData("trace-precedents", 1)]
    [InlineData("trace-precedents", 2)]
    [InlineData("trace-precedents", 4)]
    [InlineData("trace-dependents", 1)]
    [InlineData("trace-dependents", 2)]
    [InlineData("trace-dependents", 4)]
    public void Limits_NodeBudgetIncludesRootsAndNeverReturnsDanglingEdges(string action, int maxNodes)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[2]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1:D1",
            [["=A1*2", "=B1*2", "=C1*2"]]).Success);
        bool precedents = action == "trace-precedents";
        var response = _fixture.Send($"range.{action}", new
        {
            sheetName,
            rangeAddress = precedents ? "D1" : "A1",
            maxDepth = 20,
            maxNodes
        });
        using var result = JsonDocument.Parse(response.Result!);
        var addresses = Enumerable.Range(0, maxNodes)
            .Select(depth => $"${(char)((precedents ? 'D' : 'A') + (precedents ? -depth : depth))}$1").ToArray();
        AssertGraph(result, [addresses[0]],
            addresses.Zip(addresses.Skip(1), (from, to) => (from, to)).ToArray());
        var truncation = result.RootElement.GetProperty("truncation");
        Assert.Equal(maxNodes < 4, truncation.GetProperty("nodeLimitReached").GetBoolean());
        Assert.Equal(maxNodes < 4, truncation.GetProperty("any").GetBoolean());
        Assert.False(truncation.GetProperty("startingCellsOmitted").GetBoolean());
        Assert.False(truncation.GetProperty("depthLimitReached").GetBoolean());
        if (maxNodes < 4)
        {
            var boundary = Assert.Single(result.RootElement.GetProperty("boundaries").EnumerateArray());
            Assert.Equal(addresses[^1], boundary.GetProperty("address").GetString());
            Assert.Equal("max-nodes", boundary.GetProperty("reason").GetString());
        }
    }

    [Fact]
    public void Limits_OverlappingRootsKeepMinimumDepthWithoutChargingDuplicates()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[2]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1:D1",
            [["=A1*2", "=B1*2", "=B1+C1"]]).Success);
        var response = _fixture.Send("range.trace-precedents", new
        {
            sheetName,
            rangeAddress = "D1,C1,D1",
            maxDepth = 1,
            maxNodes = 3
        });
        using var result = JsonDocument.Parse(response.Result!);
        AssertGraph(result, ["$C$1", "$D$1"],
            [("$D$1", "$B$1"), ("$D$1", "$C$1"), ("$C$1", "$B$1")]);
        var nodes = result.RootElement.GetProperty("nodes").EnumerateArray().ToArray();
        Assert.Equal(1, nodes.Single(node => node.GetProperty("address").GetString() == "$B$1")
            .GetProperty("depth").GetInt32());
        Assert.All(nodes.Where(node => node.GetProperty("isRoot").GetBoolean()),
            node => Assert.Equal(0, node.GetProperty("depth").GetInt32()));
        Assert.False(result.RootElement.GetProperty("truncation").GetProperty("startingCellsOmitted").GetBoolean());
        Assert.True(result.RootElement.GetProperty("truncation").GetProperty("depthLimitReached").GetBoolean());
        Assert.True(result.RootElement.GetProperty("truncation").GetProperty("nodeLimitReached").GetBoolean());
    }

    [Theory]
    [InlineData("trace-precedents", 1)]
    [InlineData("trace-precedents", 6)]
    [InlineData("trace-precedents", 9)]
    [InlineData("trace-dependents", 1)]
    [InlineData("trace-dependents", 6)]
    [InlineData("trace-dependents", 9)]
    public void Limits_PreviewsAreFlaggedAndStoredValuesAreUnchanged(string action, int maxPreviewCharacters)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        const string formula = "=\"abcdef\"";
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1", [[formula]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "B1:D1", [[42, true, "abcdef"]]).Success);
        var response = _fixture.Send($"range.{action}", new
        {
            sheetName,
            rangeAddress = "A1:D1",
            maxDepth = 0,
            maxPreviewCharacters
        });
        using var result = JsonDocument.Parse(response.Result!);
        var nodes = result.RootElement.GetProperty("nodes").EnumerateArray().ToArray();
        Assert.Equal(formula[..Math.Min(formula.Length, maxPreviewCharacters)], nodes[0].GetProperty("formula").GetString());
        Assert.Equal("abcdef"[..Math.Min(6, maxPreviewCharacters)], nodes[0].GetProperty("value").GetString());
        Assert.Equal(maxPreviewCharacters < formula.Length, nodes[0].GetProperty("formulaTruncated").GetBoolean());
        Assert.Equal(maxPreviewCharacters < 6, nodes[0].GetProperty("valueTruncated").GetBoolean());
        Assert.Equal(42, nodes[1].GetProperty("value").GetInt32());
        Assert.True(nodes[2].GetProperty("value").GetBoolean());
        Assert.False(nodes[1].GetProperty("valueTruncated").GetBoolean());
        Assert.False(nodes[2].GetProperty("formulaTruncated").GetBoolean());
        Assert.Equal("abcdef"[..Math.Min(6, maxPreviewCharacters)], nodes[3].GetProperty("value").GetString());
        Assert.Equal(maxPreviewCharacters < 6, nodes[3].GetProperty("valueTruncated").GetBoolean());
        Assert.Equal(maxPreviewCharacters < formula.Length,
            result.RootElement.GetProperty("truncation").GetProperty("previewTextTruncated").GetBoolean());
        Assert.Equal(formula, RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheetName, "A1")).Formulas[0][0]);
        Assert.Equal("abcdef", RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "D1")).Values[0][0]);
    }

    [Fact]
    public void Limits_PreviewClippingPreservesUnicodeAndDoesNotChangeTraversalCompleteness()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [["x\uD83D\uDE00y"]]).Success);
        Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1", [["=A1"]]).Success);
        var response = _fixture.Send("range.trace-precedents", new
        {
            sheetName,
            rangeAddress = "B1",
            maxPreviewCharacters = 2
        });
        using var result = JsonDocument.Parse(response.Result!);
        AssertGraph(result, ["$B$1"], [("$B$1", "$A$1")]);
        var nodes = result.RootElement.GetProperty("nodes").EnumerateArray().ToArray();
        Assert.Equal("x", nodes[0].GetProperty("value").GetString());
        Assert.Equal("x", nodes[1].GetProperty("value").GetString());
        Assert.Equal("=A", nodes[1].GetProperty("formula").GetString());
        Assert.True(nodes[1].GetProperty("formulaTruncated").GetBoolean());
        Assert.True(result.RootElement.GetProperty("truncation").GetProperty("any").GetBoolean());
        Assert.True(result.RootElement.GetProperty("coverage").GetProperty("nativeTraversalComplete").GetBoolean());
        Assert.Empty(result.RootElement.GetProperty("boundaries").EnumerateArray());
    }

    [Theory]
    [InlineData("maxDepth", -1, "0 and 20")]
    [InlineData("maxDepth", 21, "0 and 20")]
    [InlineData("maxNodes", 0, "1 and 1000")]
    [InlineData("maxNodes", 1001, "1 and 1000")]
    [InlineData("maxPreviewCharacters", 0, "1 and 4096")]
    [InlineData("maxPreviewCharacters", 4097, "1 and 4096")]
    public async Task Limits_RejectInvalidValuesWithoutChangingCells(string parameter, int value, string bounds)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [["retained"]]).Success);
        foreach (string action in new[] { "trace-precedents", "trace-dependents" })
        {
            var response = await _fixture.SendForFailureAsync($"range.{action}", new Dictionary<string, object>
            {
                ["sheetName"] = sheetName,
                ["rangeAddress"] = "A1",
                [parameter] = value
            });
            Assert.Contains($"{parameter} must be between {bounds}", response.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal("retained", RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1")).Values[0][0]);
        }
    }

    [Theory]
    [InlineData("trace-precedents", true)]
    [InlineData("trace-precedents", false)]
    [InlineData("trace-dependents", true)]
    [InlineData("trace-dependents", false)]
    public async Task Limits_PreserveWorkbookAndViewForFullShortenedAndRejectedReports(string action, bool saved)
    {
        var calculation = _fixture.CreateCommands<ICalculationModeCommands>();
        var previous = RequireSuccess(calculation.GetSettings(_fixture.BatchToken));
        var source = _fixture.CreateTestSheet(_fixture.BatchToken);
        var active = _fixture.CreateTestSheet(_fixture.BatchToken);
        var failure = await Record.ExceptionAsync(async () =>
        {
            RequireSuccess(calculation.SetSettings(_fixture.BatchToken, CalculationMode.Manual));
            RequireSuccess(_commands.SetValues(_fixture.BatchToken, source, "A1", [[2]]));
            RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, source, "B1:C1", [["=A1*2", "=B1+1"]]));
            RequireSuccess(calculation.Calculate(_fixture.BatchToken, CalculationScope.Application));
            if (!saved)
                RequireSuccess(_commands.SetValues(_fixture.BatchToken, source, "A1", [[3]], overwritePolicy: OverwritePolicy.Allow));
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Worksheet? sheet = null;
                Excel.Range? selection = null;
                Excel.Window? window = null;
                try
                {
                    sheet = ComUtilities.FindSheet(context.Book, active);
                    Assert.NotNull(sheet);
                    sheet.Activate();
                    selection = sheet.Range["D7"];
                    selection.Select();
                    window = context.App.ActiveWindow;
                    window.ScrollRow = 3;
                    window.ScrollColumn = 2;
                    window.Zoom = 85;
                }
                finally
                {
                    ComUtilities.Release(ref window);
                    ComUtilities.Release(ref selection);
                    ComUtilities.Release(ref sheet);
                }
            });
            if (saved)
                Assert.True(_fixture.Send("session.save", new { }).Success);
            var before = ReadTraceState(source);
            Assert.Equal(saved, before.Saved);
            Assert.Equal(4d, Convert.ToDouble(
                RequireSuccess(_commands.GetValues(_fixture.BatchToken, source, "B1")).Values[0][0],
                System.Globalization.CultureInfo.InvariantCulture));
            var rangeAddress = action == "trace-precedents" ? "C1" : "A1";
            foreach (int maxDepth in PreservationDepths)
            {
                var response = _fixture.Send($"range.{action}", new
                {
                    sheetName = source,
                    rangeAddress,
                    maxDepth,
                    maxPreviewCharacters = 1
                });
                using var result = JsonDocument.Parse(response.Result!);
                Assert.True(result.RootElement.GetProperty("success").GetBoolean());
                Assert.Equal(maxDepth == 0 ? 1 : 3, result.RootElement.GetProperty("nodes").GetArrayLength());
                Assert.Equal(before, ReadTraceState(source));
            }
            var rejected = await _fixture.SendForFailureAsync($"range.{action}", new
            {
                sheetName = source,
                rangeAddress,
                maxDepth = -1
            });
            Assert.Contains("maxDepth must be between 0 and 20", rejected.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal(before, ReadTraceState(source));
        });
        var restoreFailure = Record.Exception(() =>
            RequireSuccess(calculation.SetSettings(_fixture.BatchToken, (CalculationMode)previous.ModeValue)));
        if (restoreFailure is not null)
            failure = PersistentServiceCleanupFailures.Combine(failure, restoreFailure);
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private (string Contents, string View, bool Saved) ReadTraceState(string sheetName)
    {
        string contents = JsonSerializer.Serialize(new
        {
            formulas = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheetName, "A1:C1")).Formulas,
            values = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:C1")).Values
        });
        return _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Worksheet? activeSheet = null;
            Excel.Workbook? activeBook = null;
            Excel.Range? selection = null;
            Excel.Window? window = null;
            Excel.Shapes? shapes = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                shapes = sheet.Shapes;
                activeBook = context.App.ActiveWorkbook;
                activeSheet = (Excel.Worksheet)context.App.ActiveSheet;
                selection = (Excel.Range)context.App.Selection;
                window = context.App.ActiveWindow;
                string view = JsonSerializer.Serialize(new
                {
                    workbook = activeBook.FullName,
                    sheet = activeSheet.Name,
                    selection = selection.Address,
                    window.ScrollRow,
                    window.ScrollColumn,
                    window.Zoom,
                    calculation = (int)context.App.Calculation,
                    context.App.CalculateBeforeSave,
                    context.App.Iteration,
                    context.App.MaxIterations,
                    context.App.MaxChange,
                    context.App.EnableEvents,
                    shapes = shapes.Count
                });
                return (contents, view, context.Book.Saved);
            }
            finally
            {
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref activeSheet);
                ComUtilities.Release(ref activeBook);
                ComUtilities.Release(ref sheet);
            }
        });
    }

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
        AssertGraph(result, ["$D$1"],
            [("$D$1", "$B$1"), ("$D$1", "$C$1"), ("$C$1", "$B$1"), ("$B$1", "$A$1")]);
        Assert.Equal("same-worksheet-only",
            result.RootElement.GetProperty("coverage").GetProperty("scope").GetString());
        Assert.Empty(result.RootElement.GetProperty("unresolved").EnumerateArray());
        Assert.True(result.RootElement.GetProperty("coverage").GetProperty("nativeTraversalComplete").GetBoolean());
        Assert.False(result.RootElement.GetProperty("truncation").GetProperty("any").GetBoolean());
        var limits = result.RootElement.GetProperty("limits");
        Assert.Equal(3, limits.GetProperty("maxDepth").GetInt32());
        Assert.Equal(50, limits.GetProperty("maxNodes").GetInt32());
        Assert.Equal(256, limits.GetProperty("maxPreviewCharacters").GetInt32());
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
        AssertGraph(result, ["$A$1"], [("$A$1", "$B$1"), ("$B$1", "$C$1")]);
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
        AssertGraph(result, ["$B$1"], [("$B$1", "$A$1")]);
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
        AssertGraph(result, ["$B$1", "$C$1"], [("$B$1", "$A$1"), ("$C$1", "$B$1")]);
        Assert.DoesNotContain(nodes, node => node.GetProperty("address").GetString() == "$D$1");
        Assert.Equal(2, result.RootElement.GetProperty("edges").GetArrayLength());
        var named = _fixture.Send("range.trace-precedents", new { sheetName = "", rangeAddress = name });
        using var namedResult = JsonDocument.Parse(named.Result!);
        Assert.Equal(sheetName, namedResult.RootElement.GetProperty("sheetName").GetString());
        Assert.Single(namedResult.RootElement.GetProperty("nodes").EnumerateArray());
        Assert.Empty(namedResult.RootElement.GetProperty("unresolved").EnumerateArray());
    }

    [Theory]
    [InlineData("trace-precedents", 3, 50, true)]
    [InlineData("trace-dependents", 3, 50, true)]
    [InlineData("trace-precedents", 1, 2, true)]
    [InlineData("trace-dependents", 1, 2, true)]
    [InlineData("trace-precedents", 0, 2, false)]
    [InlineData("trace-dependents", 3, 1, false)]
    public void CircularGraphs_TerminateAndIdentifyOnlyReturnedCycles(string action, int maxDepth, int maxNodes, bool complete)
    {
        var calculation = _fixture.CreateCommands<ICalculationModeCommands>();
        var previous = calculation.GetSettings(_fixture.BatchToken);
        Assert.True(previous.Success);
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var failure = Record.Exception(() =>
        {
            Assert.True(calculation.SetSettings(_fixture.BatchToken, CalculationMode.Manual).Success);
            Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1:B1", [["=B1", "=A1"]]).Success);
            var response = _fixture.Send($"range.{action}", new { sheetName, rangeAddress = "A1", maxDepth, maxNodes });
            using var result = JsonDocument.Parse(response.Result!);
            AssertGraph(result, ["$A$1"], complete ? [("$A$1", "$B$1"), ("$B$1", "$A$1")] : []);
            Assert.Equal(!complete, result.RootElement.GetProperty("truncation").GetProperty("any").GetBoolean());
            Assert.Equal(complete, result.RootElement.GetProperty("coverage").GetProperty("nativeTraversalComplete").GetBoolean());
            if (complete)
            {
                var cycle = Assert.Single(result.RootElement.GetProperty("cycles").EnumerateArray());
                Assert.Equal(["$A$1", "$B$1"], cycle.EnumerateArray().Select(cell => cell.GetString()));
            }
            else
                Assert.Empty(result.RootElement.GetProperty("cycles").EnumerateArray());
        });
        var clearFailure = Record.Exception(() =>
            RequireSuccess(_commands.ClearContents(_fixture.BatchToken, sheetName, "A1:B1")));
        if (clearFailure is not null)
            failure = PersistentServiceCleanupFailures.Combine(failure, clearFailure);
        var restoreFailure = Record.Exception(() =>
            RequireSuccess(calculation.SetSettings(_fixture.BatchToken, (CalculationMode)previous.ModeValue)));
        if (restoreFailure is not null)
            failure = PersistentServiceCleanupFailures.Combine(failure, restoreFailure);
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
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
        AssertGraph(result, ["$" + rangeAddress[0] + "$1"],
            action == "trace-precedents" ? [("$B$1", "$A$1")] : [("$A$1", "$B$1")]);
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
        AssertGraph(result, ["$C$1"], [("$C$1", "$A$2"), ("$C$1", "$A$3")]);
        Assert.Empty(result.RootElement.GetProperty("unresolved").EnumerateArray());
    }

    private static void AssertGraph(JsonDocument result, string[] expectedRoots,
        (string From, string To)[] expectedEdges)
    {
        var nodes = result.RootElement.GetProperty("nodes").EnumerateArray().ToArray();
        Assert.Equal(expectedRoots.Order(StringComparer.Ordinal), nodes
            .Where(node => node.GetProperty("isRoot").GetBoolean())
            .Select(node => node.GetProperty("address").GetString()).Order(StringComparer.Ordinal));
        Assert.Equal(expectedEdges.OrderBy(edge => edge.From, StringComparer.Ordinal)
            .ThenBy(edge => edge.To, StringComparer.Ordinal),
            result.RootElement.GetProperty("edges").EnumerateArray()
                .Select(edge => (From: edge.GetProperty("fromAddress").GetString()!,
                    To: edge.GetProperty("toAddress").GetString()!))
                .OrderBy(edge => edge.From, StringComparer.Ordinal).ThenBy(edge => edge.To, StringComparer.Ordinal));
        Assert.Equal(expectedEdges.SelectMany(edge => new[] { edge.From, edge.To })
            .Concat(expectedRoots).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal),
            nodes.Select(node => node.GetProperty("address").GetString()).Order(StringComparer.Ordinal));
    }
}
