using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Protection")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceFineProtectionTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private static readonly int[][] SingleThirteen = [[13]];

    [Fact]
    public void SheetProtection_ReadsActualPermissionsAndRuntimeOnlyMode()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var failure = Record.Exception(() =>
        {
            _fixture.Send("worksheetstyle.set-protection", new
            {
                sheetName,
                isProtected = true,
                options = new
                {
                    allowFiltering = true,
                    allowFormattingRows = true,
                    userInterfaceOnly = true
                }
            });
            var response = _fixture.Send("worksheetstyle.get-protection", new { sheetName });
            using var result = JsonDocument.Parse(response.Result!);
            Assert.True(result.RootElement.GetProperty("isProtected").GetBoolean());
            Assert.True(result.RootElement.GetProperty("userInterfaceOnly").GetBoolean());
            var permissions = result.RootElement.GetProperty("permissions");
            Assert.True(permissions.GetProperty("allowFiltering").GetBoolean());
            Assert.True(permissions.GetProperty("allowFormattingRows").GetBoolean());
            Assert.False(permissions.GetProperty("allowSorting").GetBoolean());
        });
        UnprotectAndThrowFailure(sheetName, null, failure);
    }

    [Fact]
    public void CellProtection_ExactScopePreservesGapsAndReadsEveryCell()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.Send("rangelink.set-cell-protection", new
        {
            sheetName,
            rangeAddress = "A1,A3",
            locked = false,
            formulaHidden = true
        });
        var response = _fixture.Send("rangelink.get-cell-protection", new
        {
            sheetName,
            rangeAddress = "A1:A3"
        });
        using var result = JsonDocument.Parse(response.Result!);
        var cells = result.RootElement.GetProperty("cells");
        Assert.Equal(3, cells.GetArrayLength());
        Assert.False(cells[0].GetProperty("locked").GetBoolean());
        Assert.True(cells[0].GetProperty("formulaHidden").GetBoolean());
        Assert.True(cells[1].GetProperty("locked").GetBoolean());
        Assert.False(cells[1].GetProperty("formulaHidden").GetBoolean());
        Assert.False(cells[2].GetProperty("locked").GetBoolean());
        Assert.True(cells[2].GetProperty("formulaHidden").GetBoolean());
    }

    [Fact]
    public async Task SheetProtection_AllowedRowFormattingWorksButOtherChangesAreRejected()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var before = _fixture.Send("rangeformat.get-visibility",
            new { sheetName, rangeAddress = "A2:A3", axis = "columns" });
        var failure = await Record.ExceptionAsync(async () =>
        {
            _fixture.Send("worksheetstyle.set-protection", new
            {
                sheetName,
                isProtected = true,
                options = new { allowFormattingRows = true }
            });
            _fixture.Send("rangeformat.set-row-height", new { sheetName, rangeAddress = "A2:A3", rowHeight = 27d });
            var read = _fixture.Send("rangeformat.get-visibility", new { sheetName, rangeAddress = "A2:A3", axis = "rows" });
            using var result = JsonDocument.Parse(read.Result!);
            Assert.All(result.RootElement.GetProperty("items").EnumerateArray(),
                row => Assert.Equal(27d, row.GetProperty("size").GetDouble()));
            var forbidden = await _fixture.SendForFailureAsync("rangeformat.set-column-width", new
            {
                sheetName,
                rangeAddress = "A2:A3",
                columnWidth = 20d
            });
            Assert.False(forbidden.Success);
            Assert.Equal("rangeformat.set-column-width", forbidden.Command);
            Assert.Equal(before.Result, _fixture.Send("rangeformat.get-visibility",
                new { sheetName, rangeAddress = "A2:A3", axis = "columns" }).Result);
            var retained = _fixture.Send("worksheetstyle.get-protection", new { sheetName });
            using var state = JsonDocument.Parse(retained.Result!);
            Assert.True(state.RootElement.GetProperty("protectContents").GetBoolean());
        });
        UnprotectAndThrowFailure(sheetName, null, failure);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task SheetProtection_RuntimeAutomationPermissionMatchesNativeBehavior(bool uiOnly)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var failure = await Record.ExceptionAsync(async () =>
        {
            _fixture.Send("worksheetstyle.set-protection", new
            {
                sheetName,
                isProtected = true,
                options = new { userInterfaceOnly = uiOnly }
            });
            if (uiOnly)
            {
                Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[13]]).Success);
                var values = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
                Assert.True(values.Success);
                Assert.Equal(13d, Convert.ToDouble(values.Values[0][0],
                    System.Globalization.CultureInfo.InvariantCulture));
            }
            else
            {
                var forbidden = await _fixture.SendForFailureAsync("range.set-values", new
                {
                    sheetName,
                    rangeAddress = "A1",
                    values = SingleThirteen
                });
                Assert.False(forbidden.Success);
                var values = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
                Assert.True(values.Success);
                Assert.Null(values.Values[0][0]);
            }
        });
        UnprotectAndThrowFailure(sheetName, null, failure);
    }

    [Fact]
    public void CellProtection_ChangingOneFlagPreservesTheOtherAndNamedScope()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Protection_{Guid.NewGuid():N}";
        _fixture.Send("namedrange.create", new { name, reference = $"'{sheetName}'!$A$1:$A$3" });
        _fixture.RegisterNamedRangeForCleanup(name);
        _fixture.Send("rangelink.set-cell-protection", new
        {
            sheetName = string.Empty,
            rangeAddress = name,
            formulaHidden = true
        });
        var read = _fixture.Send("rangelink.get-cell-protection", new
        {
            sheetName,
            rangeAddress = "A1:A3,A2"
        });
        using var result = JsonDocument.Parse(read.Result!);
        Assert.Equal(3, result.RootElement.GetProperty("cells").GetArrayLength());
        Assert.All(result.RootElement.GetProperty("cells").EnumerateArray(), cell =>
        {
            Assert.True(cell.GetProperty("locked").GetBoolean());
            Assert.True(cell.GetProperty("formulaHidden").GetBoolean());
        });
    }

    [Fact]
    public async Task SheetProtection_UnknownOptionsDoNotMutateProtection()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var failed = await _fixture.SendForFailureAsync("worksheetstyle.set-protection", new
        {
            sheetName,
            isProtected = true,
            options = new { allowTypo = true }
        });
        Assert.False(failed.Success);
        var read = _fixture.Send("worksheetstyle.get-protection", new { sheetName });
        using var result = JsonDocument.Parse(read.Result!);
        Assert.False(result.RootElement.GetProperty("isProtected").GetBoolean());
    }

    [Fact]
    public void SheetProtection_ReadsEveryNativePermissionAndSelectionRestriction()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var failure = Record.Exception(() =>
        {
            _fixture.Send("worksheetstyle.set-protection", new
            {
                sheetName,
                isProtected = true,
                options = new
                {
                    userInterfaceOnly = true,
                    allowFormattingCells = true,
                    allowFormattingColumns = true,
                    allowFormattingRows = true,
                    allowInsertingColumns = true,
                    allowInsertingRows = true,
                    allowInsertingHyperlinks = true,
                    allowDeletingColumns = true,
                    allowDeletingRows = true,
                    allowSorting = true,
                    allowFiltering = true,
                    allowUsingPivotTables = true,
                    selection = "UnlockedCells"
                }
            });
            var response = _fixture.Send("worksheetstyle.get-protection", new { sheetName });
            using var result = JsonDocument.Parse(response.Result!);
            var permissions = result.RootElement.GetProperty("permissions");
            foreach (var permission in permissions.EnumerateObject())
            {
                if (permission.Name == "selection")
                    Assert.Equal("UnlockedCells", permission.Value.GetString());
                else
                    Assert.True(permission.Value.GetBoolean(), permission.Name);
            }
            Assert.Equal(16, permissions.EnumerateObject().Count());
        });
        UnprotectAndThrowFailure(sheetName, null, failure);
    }

    [Fact]
    public void SheetProtection_DrawingOnlyProtectionIsNotReportedAsUnprotected()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var failure = Record.Exception(() =>
        {
            _fixture.Send("worksheetstyle.set-protection", new
            {
                sheetName,
                isProtected = true,
                options = new { contents = false, scenarios = false }
            });
            var read = _fixture.Send("worksheetstyle.get-protection", new { sheetName });
            using var result = JsonDocument.Parse(read.Result!);
            Assert.True(result.RootElement.GetProperty("isProtected").GetBoolean());
            Assert.True(result.RootElement.GetProperty("protectDrawingObjects").GetBoolean());
            Assert.False(result.RootElement.GetProperty("protectContents").GetBoolean());
            Assert.False(result.RootElement.GetProperty("protectScenarios").GetBoolean());
        });
        UnprotectAndThrowFailure(sheetName, null, failure);
    }

    [Fact]
    public async Task SheetProtection_WrongPasswordFailsWithoutBeingReturnedOrRemovingProtection()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        const string password = "Synthetic protection test";
        var failure = await Record.ExceptionAsync(async () =>
        {
            var set = _fixture.Send("worksheetstyle.set-protection", new
            {
                sheetName,
                isProtected = true,
                password,
                options = new { allowFiltering = true, allowFormattingRows = true, selection = "UnlockedCells" }
            });
            Assert.DoesNotContain(password, set.Result ?? "", StringComparison.Ordinal);
            var before = _fixture.Send("worksheetstyle.get-protection", new { sheetName });
            var failed = await _fixture.SendForFailureAsync("worksheetstyle.set-protection", new
            {
                sheetName,
                isProtected = false,
                password = "Wrong synthetic password"
            });
            Assert.False(failed.Success);
            var read = _fixture.Send("worksheetstyle.get-protection", new { sheetName });
            Assert.Equal(before.Result, read.Result);
            Assert.DoesNotContain(password, read.Result!, StringComparison.Ordinal);
            using var result = JsonDocument.Parse(read.Result!);
            Assert.True(result.RootElement.GetProperty("isProtected").GetBoolean());
            _fixture.Send("worksheetstyle.set-protection", new { sheetName, isProtected = false, password });
            var recovered = _fixture.Send("worksheetstyle.get-protection", new { sheetName });
            using var recovery = JsonDocument.Parse(recovered.Result!);
            Assert.False(recovery.RootElement.GetProperty("isProtected").GetBoolean());
        });
        UnprotectAndThrowFailure(sheetName, password, failure);
    }

    private void UnprotectAndThrowFailure(string sheetName, string? password, Exception? failure)
    {
        var cleanup = Record.Exception(() =>
            _fixture.Send("worksheetstyle.set-protection", new { sheetName, isProtected = false, password }));
        if (cleanup is not null)
            failure = PersistentServiceCleanupFailures.Combine(failure, cleanup);
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }
}
