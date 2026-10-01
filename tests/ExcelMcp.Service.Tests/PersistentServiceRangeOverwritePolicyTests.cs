using System.Globalization;
using System.Text.RegularExpressions;
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
public sealed class PersistentServiceRangeOverwritePolicyTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Theory]
    [InlineData("set-values")]
    [InlineData("set-formulas")]
    public async Task DefaultPolicy_OccupiedDestination_RejectsWithoutPartialWrite(string action)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "B1", [["Original"]]).Success);

        object args = action == "set-values"
            ? new { sheetName, rangeAddress = "A1:B1", values = new[] { new object[] { 1, 2 } } }
            : new { sheetName, rangeAddress = "A1:B1", formulas = new[] { new[] { "=1", "=2" } } };
        var response = await _fixture.SendForFailureAsync($"range.{action}", args);

        Assert.Equal("Conflict", response.ErrorCategory);
        Assert.Contains("$B$1", response.ErrorMessage);
        var read = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Null(read.Values[0][0]);
        Assert.Equal("Original", read.Values[0][1]);
    }

    [Theory]
    [InlineData("copy")]
    [InlineData("copy-values")]
    [InlineData("copy-formulas")]
    public async Task DefaultPolicy_CopyAnchor_ChecksExpandedDestination(string action)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B2", [[1, 2], [3, 4]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "E2", [["Original"]]).Success);

        var response = await _fixture.SendForFailureAsync(
            $"range.{action}",
            new { sourceSheet = sheetName, sourceRange = "A1:B2", targetSheet = sheetName, targetRange = "D1" });

        Assert.Equal("Conflict", response.ErrorCategory);
        Assert.Contains("$E$2", response.ErrorMessage);
        var read = _commands.GetValues(batch, sheetName, "D1:E2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Null(read.Values[0][0]);
        Assert.Null(read.Values[0][1]);
        Assert.Null(read.Values[1][0]);
        Assert.Equal("Original", read.Values[1][1]);
    }

    [Theory]
    [InlineData("value", "Original")]
    [InlineData("value", " ")]
    [InlineData("value", 0)]
    [InlineData("value", false)]
    [InlineData("formula", "=\"\"")]
    [InlineData("formula", "=1/0")]
    public async Task RejectNonempty_StoredContent_IsOccupied(string kind, object content)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var setup = kind == "formula"
            ? _commands.SetFormulas(batch, sheetName, "A1", [[(string)content]])
            : _commands.SetValues(batch, sheetName, "A1", [[content]]);
        Assert.True(setup.Success, setup.ErrorMessage);
        var before = _commands.GetFormulas(batch, sheetName, "A1");
        Assert.True(before.Success, before.ErrorMessage);

        var response = await _fixture.SendForFailureAsync("range.set-values",
            new { sheetName, rangeAddress = "A1", values = SingleValue("Updated"), overwritePolicy = "reject-nonempty" });

        Assert.Equal("Conflict", response.ErrorCategory);
        Assert.Contains(sheetName, response.ErrorMessage);
        Assert.Contains("$A$1", response.ErrorMessage);
        Assert.DoesNotContain("Original", response.ErrorMessage);
        var after = _commands.GetFormulas(batch, sheetName, "A1");
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(before.Formulas[0][0], after.Formulas[0][0]);
        Assert.Equal(before.Values[0][0], after.Values[0][0]);
    }

    [Theory]
    [InlineData("set-values")]
    [InlineData("set-formulas")]
    public void Allow_IntentionalReplacement_Succeeds(string action)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["Original"]]).Success);

        var result = action == "set-values"
            ? _commands.SetValues(batch, sheetName, "A1", [[42]], overwritePolicy: OverwritePolicy.Allow)
            : _commands.SetFormulas(batch, sheetName, "A1", [["=42"]], overwritePolicy: OverwritePolicy.Allow);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        var read = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(42.0, Number(read.Values[0][0]));
    }

    [Theory]
    [InlineData("copy")]
    [InlineData("copy-values")]
    [InlineData("copy-formulas")]
    public void Allow_CopyReplacesContentAndSourceBlanks(string action)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B1", [[42, null]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "D1:E1", [["Old", "Old"]]).Success);

        var response = _fixture.Send($"range.{action}", new
        {
            sourceSheet = sheetName,
            sourceRange = "A1:B1",
            targetSheet = sheetName,
            targetRange = "D1",
            overwritePolicy = "allow"
        });

        Assert.True(response.Success, response.ErrorMessage);
        var read = _commands.GetValues(batch, sheetName, "D1:E1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(42.0, Number(read.Values[0][0]));
        Assert.Null(read.Values[0][1]);
    }

    [Fact]
    public async Task SameValueAndIncomingBlank_StillRejectOccupiedCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[42]]).Success);
        foreach (object? incoming in new object?[] { 42, null, string.Empty })
        {
            var response = await _fixture.SendForFailureAsync("range.set-values",
                new { sheetName, rangeAddress = "A1", values = new[] { new[] { incoming } } });
            Assert.Equal("Conflict", response.ErrorCategory);
        }

        var read = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(42.0, Number(read.Values[0][0]));
    }

    [Fact]
    public async Task FormulaAutoRouting_PreservesPolicy()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[42]]).Success);
        var rejected = await _fixture.SendForFailureAsync("range.set-values",
            new { sheetName, rangeAddress = "A1", values = SingleValue("=21*2") });
        Assert.Equal("Conflict", rejected.ErrorCategory);

        var allowed = _commands.SetValues(batch, sheetName, "A1", [["=21*2"]], overwritePolicy: OverwritePolicy.Allow);
        Assert.True(allowed.Success, allowed.ErrorMessage);
        var read = _commands.GetFormulas(batch, sheetName, "A1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal("=21*2", read.Formulas[0][0]);
        Assert.Equal(42.0, Number(read.Values[0][0]));
    }

    [Fact]
    public void DefaultPolicy_FormattingOnlyAndStoredEmptyText_UsesActualContent()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "A1", "0.00").Success);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[42]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "B1", [[""]]).Success);
        var empty = _commands.GetValues(batch, sheetName, "B1");
        Assert.True(empty.Success, empty.ErrorMessage);
        Assert.Null(empty.Values[0][0]);
        Assert.True(_commands.SetValues(batch, sheetName, "B1", [[1]]).Success);
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "A1", "0").Success);
        var read = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(42.0, Number(read.Values[0][0]));
    }

    [Theory]
    [InlineData("copy")]
    [InlineData("copy-values")]
    [InlineData("copy-formulas")]
    public async Task ProtectedCopy_RepetitionChecksAllCellsAndOutsideCellsAreIgnored(string action)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B1", [[1, 2]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "G2", [[99]]).Success);
        var args = new { sourceSheet = sheetName, sourceRange = "A1:B1", targetSheet = sheetName, targetRange = "D1:G2" };

        var rejected = await _fixture.SendForFailureAsync($"range.{action}", args);
        Assert.Equal("Conflict", rejected.ErrorCategory);
        Assert.Contains("$G$2", rejected.ErrorMessage);
        Assert.True(_commands.ClearContents(batch, sheetName, "G2").Success);
        Assert.True(_commands.SetValues(batch, sheetName, "H2", [[99]]).Success);
        var copied = _fixture.Send($"range.{action}", args);
        Assert.True(copied.Success, copied.ErrorMessage);
        var read = _commands.GetValues(batch, sheetName, "D1:H2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(new double?[] { 1, 2, 1, 2, null }, read.Values[0].Select(Number));
        Assert.Equal(new double?[] { 1, 2, 1, 2, 99 }, read.Values[1].Select(Number));
    }

    [Theory]
    [InlineData("copy")]
    [InlineData("copy-values")]
    [InlineData("copy-formulas")]
    public async Task ProtectedCopy_SourceBlankWouldClearContent_Rejects(string action)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[42]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "E1", [["Keep"]]).Success);
        var rejected = await _fixture.SendForFailureAsync($"range.{action}", new
        {
            sourceSheet = sheetName,
            sourceRange = "A1:B1",
            targetSheet = sheetName,
            targetRange = "D1"
        });
        Assert.Equal("Conflict", rejected.ErrorCategory);
        var read = _commands.GetValues(batch, sheetName, "D1:E1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Null(read.Values[0][0]);
        Assert.Equal("Keep", read.Values[0][1]);
    }

    [Theory]
    [InlineData("D1:F1", "InvalidInput")]
    [InlineData("XFD1048576", "InvalidInput")]
    [InlineData("D1,D3", "ComInterop")]
    public async Task ProtectedCopy_UninspectableDestination_StopsWithoutMutation(string targetRange, string expectedCategory)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B1", [[1, 2]]).Success);
        var rejected = await _fixture.SendForFailureAsync("range.copy", new
        {
            sourceSheet = sheetName,
            sourceRange = "A1:B1",
            targetSheet = sheetName,
            targetRange
        });
        Assert.Equal(expectedCategory, rejected.ErrorCategory);
        var read = _commands.GetValues(batch, sheetName, "D1:F3");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.All(read.Values, row => Assert.All(row, Assert.Null));
        var edge = _commands.GetValues(batch, sheetName, "XFD1048576");
        Assert.True(edge.Success, edge.ErrorMessage);
        Assert.Null(edge.Values[0][0]);
    }

    [Fact]
    public async Task OverlappingCopy_RejectsWithoutChangingSource()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B1", [[1, 2]]).Success);
        var rejected = await _fixture.SendForFailureAsync("range.copy-values", new
        {
            sourceSheet = sheetName,
            sourceRange = "A1:B1",
            targetSheet = sheetName,
            targetRange = "B1"
        });
        Assert.Equal("Conflict", rejected.ErrorCategory);
        var read = _commands.GetValues(batch, sheetName, "A1:C1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(new double?[] { 1, 2, null }, read.Values[0].Select(Number));
    }

    [Theory]
    [InlineData(10, false)]
    [InlineData(11, true)]
    public async Task ConflictExamples_AreBoundedAndDoNotExposeContent(int count, bool truncated)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var original = Enumerable.Range(1, count).Select(_ => new List<object?> { "Private content" }).ToList();
        Assert.True(_commands.SetValues(batch, sheetName, $"A1:A{count}", original).Success);
        var rejected = await _fixture.SendForFailureAsync("range.set-values",
            new { sheetName, rangeAddress = $"A1:A{count}", values = original });
        Assert.Equal("Conflict", rejected.ErrorCategory);
        Assert.Equal(10, Regex.Matches(rejected.ErrorMessage!, @"\$A\$\d+").Count);
        Assert.Equal(truncated, rejected.ErrorMessage!.Contains("additional conflicts", StringComparison.Ordinal));
        Assert.DoesNotContain("Private content", rejected.ErrorMessage);
        var read = _commands.GetValues(batch, sheetName, $"A1:A{count}");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.All(read.Values, row => Assert.Equal("Private content", row[0]));
    }

    [Fact]
    public async Task MultiBlockInspection_ChecksLastBlockBeforeAnyWrite()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A16385", [[42]]).Success);
        var payload = Enumerable.Range(1, 16_385).Select(_ => new List<object?> { 1 }).ToList();
        var rejected = await _fixture.SendForFailureAsync("range.set-values",
            new { sheetName, rangeAddress = "A1:A16385", values = payload });
        Assert.Equal("Conflict", rejected.ErrorCategory);
        Assert.Contains("$A$16385", rejected.ErrorMessage);
        var first = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(first.Success, first.ErrorMessage);
        Assert.Null(first.Values[0][0]);
    }

    [Theory]
    [InlineData("unknown")]
    [InlineData("1")]
    [InlineData(" ")]
    public async Task InvalidPolicy_RejectsBeforeWriting(string overwritePolicy)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var rejected = await _fixture.SendForFailureAsync("range.set-values", new
        {
            sheetName,
            rangeAddress = "A1",
            values = SingleValue(1),
            overwritePolicy
        });
        Assert.Equal("InvalidInput", rejected.ErrorCategory);
        var read = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Null(read.Values[0][0]);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public async Task EmptyOrNullPolicy_NeverDisablesProtection(string? overwritePolicy)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[42]]).Success);
        var rejected = await _fixture.SendForFailureAsync("range.set-values", new
        {
            sheetName,
            rangeAddress = "A1",
            values = SingleValue(1),
            overwritePolicy
        });
        Assert.Equal("Conflict", rejected.ErrorCategory);
    }

    [Fact]
    public async Task FileInput_UsesSamePolicy()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        string file = _fixture.CreateInputFile(".json", "[[42]]");
        Assert.True(_commands.SetValues(batch, sheetName, "A1", valuesFile: file).Success);
        var rejected = await _fixture.SendForFailureAsync("range.set-formulas", new
        {
            sheetName,
            rangeAddress = "A1",
            formulasFile = _fixture.CreateInputFile(".json", """[["=42"]]""")
        });
        Assert.Equal("Conflict", rejected.ErrorCategory);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", valuesFile: file, overwritePolicy: OverwritePolicy.Allow).Success);
    }

    [Fact]
    public async Task NamedDestination_ReportsResolvedSheetAndAddress()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        string name = $"Guarded_{Guid.NewGuid():N}";
        Assert.True(_fixture.Send("namedrange.create", new { name, reference = $"='{sheetName}'!$A$1" }).Success);
        _fixture.RegisterNamedRangeForCleanup(name);
        Assert.True(_commands.SetValues(batch, "", name, [[42]]).Success);
        var rejected = await _fixture.SendForFailureAsync("range.set-values",
            new { sheetName = "", rangeAddress = name, values = SingleValue(1) });
        Assert.Equal("Conflict", rejected.ErrorCategory);
        Assert.Contains(sheetName, rejected.ErrorMessage);
        Assert.Contains("$A$1", rejected.ErrorMessage);
    }

    [Fact]
    public async Task MergedAnchor_ProtectsContentAndCopyRejectsAmbiguousMergeEffects()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.MergeCells(batch, sheetName, "A1:B1").Success);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[42]]).Success);
        var rejected = await _fixture.SendForFailureAsync("range.set-values",
            new { sheetName, rangeAddress = "A1", values = SingleValue(1) });
        Assert.Equal("Conflict", rejected.ErrorCategory);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[43]], overwritePolicy: OverwritePolicy.Allow).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "D1", [[1]]).Success);
        var copied = await _fixture.SendForFailureAsync("range.copy", new
        {
            sourceSheet = sheetName,
            sourceRange = "D1",
            targetSheet = sheetName,
            targetRange = "A1"
        });
        Assert.Equal("Conflict", copied.ErrorCategory);
        Assert.Contains("merged", copied.ErrorMessage);
        var read = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(43.0, Number(read.Values[0][0]));
    }

    [Theory]
    [InlineData("allow")]
    [InlineData("reject-nonempty")]
    public async Task ProtectedSheet_WriteFailureRemainsAnError(string overwritePolicy)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        void Protect(bool enabled) => _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
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
            var rejected = await _fixture.SendForFailureAsync("range.set-values", new
            {
                sheetName,
                rangeAddress = "A1",
                values = SingleValue(1),
                overwritePolicy
            });
            Assert.False(rejected.Success);
            Assert.False(string.IsNullOrWhiteSpace(rejected.ErrorMessage));
            var read = _commands.GetValues(batch, sheetName, "A1");
            Assert.True(read.Success, read.ErrorMessage);
            Assert.Null(read.Values[0][0]);
        }
        finally
        {
            Protect(false);
        }
    }

    [Theory]
    [InlineData(CalculationMode.Automatic)]
    [InlineData(CalculationMode.Manual)]
    [InlineData(CalculationMode.SemiAutomatic)]
    public async Task RejectedWrite_DoesNotChangeCalculationMode(CalculationMode mode)
    {
        var calculation = _fixture.CreateCommands<ICalculationModeCommands>();
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[42]]).Success);
        var previous = calculation.GetMode(batch);
        Assert.True(previous.Success, previous.ErrorMessage);
        try
        {
            Assert.True(calculation.SetMode(batch, mode).Success);
            var rejected = await _fixture.SendForFailureAsync("range.set-values",
                new { sheetName, rangeAddress = "A1", values = SingleValue(1) });
            Assert.Equal("Conflict", rejected.ErrorCategory);
            var retained = calculation.GetMode(batch);
            Assert.True(retained.Success, retained.ErrorMessage);
            Assert.Equal((int)mode, retained.ModeValue);
        }
        finally
        {
            Assert.True(calculation.SetMode(batch, (CalculationMode)previous.ModeValue).Success);
        }
    }

    [Theory]
    [InlineData("set-values")]
    [InlineData("set-formulas")]
    public async Task PayloadRowMismatch_RejectsBeforeMutation(string action)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        object args = action == "set-values"
            ? new { sheetName, rangeAddress = "A1:A2", values = SingleValue(1) }
            : new { sheetName, rangeAddress = "A1:A2", formulas = new List<List<string>> { new() { "=1" } } };
        var rejected = await _fixture.SendForFailureAsync($"range.{action}", args);
        Assert.Equal("InvalidInput", rejected.ErrorCategory);
        Assert.Contains("row count", rejected.ErrorMessage);
        var read = _commands.GetValues(batch, sheetName, "A1:A2");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.All(read.Values, row => Assert.Null(row[0]));
    }

    [Fact]
    public async Task RejectedCopy_PreservesDestinationFormatting()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B1", [[1, 2]]).Success);
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "A1:B1", "0.00").Success);
        Assert.True(_commands.SetValues(batch, sheetName, "E1", [[42]]).Success);
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "D1:E1", "0%").Success);
        var rejected = await _fixture.SendForFailureAsync("range.copy", new
        {
            sourceSheet = sheetName,
            sourceRange = "A1:B1",
            targetSheet = sheetName,
            targetRange = "D1"
        });
        Assert.Equal("Conflict", rejected.ErrorCategory);
        var formats = _commands.GetNumberFormats(batch, sheetName, "D1:E1");
        Assert.True(formats.Success, formats.ErrorMessage);
        Assert.All(formats.Formats, row => Assert.All(row, format => Assert.Equal("0%", format)));
    }

    private static List<List<object?>> SingleValue(object? value) => [[value]];

    private static double? Number(object? value) =>
        value is null ? null : Convert.ToDouble(value, CultureInfo.InvariantCulture);
}
