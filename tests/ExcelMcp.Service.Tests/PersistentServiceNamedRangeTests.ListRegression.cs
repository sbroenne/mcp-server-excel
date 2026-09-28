using Sbroenne.ExcelMcp.ComInterop;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceNamedRangeTests
{
    [Fact]
    public void List_HiddenUserDefinedName_DoesNotExposeHiddenNameByDefault()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var name = CreateUniqueNamedRangeName();
        _fixture.RegisterNamedRangeForCleanup(name);
        AddName(_fixture, name, $"{sheetName}!$B$4", visible: false);

        var result = _parameterCommands.List(batch);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.DoesNotContain(result.NamedRanges, namedRange => namedRange.Name == name);
    }

    [Fact]
    public void List_LargeNamedRange_OmitsValuePreview()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var name = CreateUniqueNamedRangeName();
        _fixture.RegisterNamedRangeForCleanup(name);
        AddName(_fixture, name, $"{sheetName}!$A$1:$A$10001");

        var result = _parameterCommands.List(batch);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        var listedRange = Assert.Single(result.NamedRanges, namedRange => namedRange.Name == name);
        Assert.Equal("RangeTooLarge", listedRange.ValueType);
        Assert.Null(listedRange.Value);
        Assert.Equal(10001, listedRange.CellCount);
        Assert.Contains("exceeds", listedRange.ValueOmittedReason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void List_MultipleLargeHiddenExternalDataNames_ReturnsOnlyVisibleUserNames()
    {
        var batch = _fixture.BatchToken;
        var sourceSheet = _fixture.CreateTestSheet(batch);
        var usersSheet = _fixture.CreateTestSheet(batch);
        var notificationsSheet = _fixture.CreateTestSheet(batch);
        var visibleName = CreateUniqueNamedRangeName();
        _fixture.RegisterNamedRangeForCleanup(visibleName);

        SetCellValue(_fixture, sourceSheet, "$B$4", "C:\\Data");
        AddName(_fixture, visibleName, $"{sourceSheet}!$B$4");
        AddSheetScopedName(_fixture, usersSheet, "ExternalData_1", $"{usersSheet}!$A$6:$AH$19132");
        AddSheetScopedName(
            _fixture,
            notificationsSheet,
            "ExternalData_1",
            $"{notificationsSheet}!$A$6:$P$28365");

        var result = _parameterCommands.List(batch);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        var listedRange = Assert.Single(result.NamedRanges);
        Assert.Equal(visibleName, listedRange.Name);
        Assert.DoesNotContain(
            result.NamedRanges,
            namedRange => namedRange.Name.Contains("ExternalData_1", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void List_JapaneseSheetsWithMultipleLargeHiddenExternalDataNames_ReturnsVisibleUserName()
    {
        var batch = _fixture.BatchToken;
        var suffix = Guid.NewGuid().ToString("N")[..6];
        var settingsSheet = _fixture.CreateNamedTestSheet(batch, $"PQ_設定_{suffix}");
        var usersSheet = _fixture.CreateNamedTestSheet(batch, $"ユーザーテーブル_{suffix}");
        var notificationsSheet = _fixture.CreateNamedTestSheet(batch, $"通知テーブル_{suffix}");
        var visibleName = CreateUniqueNamedRangeName();
        _fixture.RegisterNamedRangeForCleanup(visibleName);

        SetCellValue(_fixture, settingsSheet, "$B$4", "C:\\Data");
        AddName(_fixture, visibleName, $"{settingsSheet}!$B$4");
        AddSheetScopedName(_fixture, usersSheet, "ExternalData_1", $"{usersSheet}!$A$6:$AH$19132");
        AddSheetScopedName(
            _fixture,
            notificationsSheet,
            "ExternalData_1",
            $"{notificationsSheet}!$A$6:$P$28365");

        var result = _parameterCommands.List(batch);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        var listedRange = Assert.Single(result.NamedRanges);
        Assert.Equal(visibleName, listedRange.Name);
        Assert.Contains("PQ_設定", listedRange.RefersTo, StringComparison.Ordinal);
        Assert.Equal("C:\\Data", listedRange.Value);
        Assert.DoesNotContain(
            result.NamedRanges,
            namedRange => namedRange.Name.Contains("ExternalData_1", StringComparison.OrdinalIgnoreCase));
    }

    private static void AddName(
        PersistentServiceWorkbookTestScope scope,
        string name,
        string reference,
        bool visible = true) =>
        scope.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? names = null;
            dynamic? nameObject = null;
            try
            {
                names = ctx.Book.Names;
                nameObject = names.Add(name, $"={reference.TrimStart('=')}");
                nameObject.Visible = visible;
            }
            finally
            {
                ComUtilities.Release(ref nameObject);
                ComUtilities.Release(ref names);
            }
        });

    private static void AddSheetScopedName(
        PersistentServiceWorkbookTestScope scope,
        string sheetName,
        string name,
        string reference) =>
        scope.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? names = null;
            dynamic? nameObject = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                names = sheet.Names;
                nameObject = names.Add(name, $"={reference.TrimStart('=')}");
                nameObject.Visible = false;
            }
            finally
            {
                ComUtilities.Release(ref nameObject);
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref sheet);
            }
        });

    private static void SetCellValue(
        PersistentServiceWorkbookTestScope scope,
        string sheetName,
        string rangeAddress,
        string value) =>
        scope.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                range = sheet.Range[rangeAddress];
                range.Value2 = value;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
}
