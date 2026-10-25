using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceNamedRangeTests
{
    [Fact]
    public void List_MultiAreaNamedRange_ReportsOmittedPreviewWithoutChangingCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var name = CreateUniqueNamedRangeName();
        Assert.True(_commands.SetValues(batch, sheetName, "A1:C1",
            [["First", "Untouched", "Last"]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? first = null;
            Excel.Range? last = null;
            Excel.Range? union = null;
            Excel.Areas? areas = null;
            Excel.Names? names = null;
            Excel.Name? created = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                first = sheet.Range["A1"];
                last = sheet.Range["C1"];
                union = context.App.Union(first, last);
                areas = union.Areas;
                Assert.Equal(2, areas.Count);
                names = context.Book.Names;
                created = names.Add(Name: name, RefersTo: union);
                Assert.Equal(name, created.Name);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref created);
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref areas);
                ComUtilities.Release(ref union);
                ComUtilities.Release(ref last);
                ComUtilities.Release(ref first);
                ComUtilities.Release(ref sheet);
            }
        });
        _fixture.RegisterNamedRangeForCleanup(name);

        var listed = _parameterCommands.List(batch);

        Assert.True(listed.Success, listed.ErrorMessage);
        var entry = Assert.Single(listed.NamedRanges, item => item.Name == name);
        Assert.Equal("MultiAreaRange", entry.ValueType);
        Assert.Null(entry.Value);
        Assert.Contains("multiple areas", entry.ValueOmittedReason, StringComparison.Ordinal);
        var cells = _commands.GetValues(batch, sheetName, "A1:C1");
        Assert.True(cells.Success, cells.ErrorMessage);
        Assert.Equal(["First", "Untouched", "Last"], Assert.Single(cells.Values));
    }

    [Fact]
    public void List_HiddenUserDefinedName_DoesNotExposeHiddenNameByDefault()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var name = CreateUniqueNamedRangeName();
        _fixture.RegisterNamedRangeForCleanup(name);
        AddName(_fixture, name, $"{sheetName}!$B$4", visible: false);
        var visibleName = CreateUniqueNamedRangeName();
        _fixture.RegisterNamedRangeForCleanup(visibleName);
        SetCellValue(_fixture, sheetName, "$C$4", "Visible value");
        AddName(_fixture, visibleName, $"{sheetName}!$C$4");

        var result = _parameterCommands.List(batch);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.DoesNotContain(result.NamedRanges, namedRange => namedRange.Name == name);
        var visible = Assert.Single(result.NamedRanges);
        Assert.Equal(visibleName, visible.Name);
        Assert.Equal($"={sheetName}!$C$4", visible.RefersTo);
        Assert.Equal("Visible value", visible.Value);
        Assert.Equal("String", visible.ValueType);
        Assert.Equal(1, visible.CellCount);
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
        Assert.Equal($"={sheetName}!$A$1:$A$10001", listedRange.RefersTo);
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
        Assert.Equal($"={sourceSheet}!$B$4", listedRange.RefersTo);
        Assert.Equal("C:\\Data", listedRange.Value);
        Assert.Equal("String", listedRange.ValueType);
        Assert.Equal(1, listedRange.CellCount);
        Assert.DoesNotContain(
            result.NamedRanges,
            namedRange => namedRange.Name.Contains("ExternalData_1", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void List_JapaneseSheetsWithMultipleLargeHiddenExternalDataNames_ReturnsVisibleUserName()
    {
        var batch = _fixture.BatchToken;
        var suffix = Guid.NewGuid().ToString("N")[..6];
        var settingsSheet = _fixture.CreateNamedTestSheet(batch, $"PQ_設定 {suffix}");
        var usersSheet = _fixture.CreateNamedTestSheet(batch, $"ユーザーテーブル_{suffix}");
        var notificationsSheet = _fixture.CreateNamedTestSheet(batch, $"通知テーブル_{suffix}");
        var visibleName = CreateUniqueNamedRangeName();
        _fixture.RegisterNamedRangeForCleanup(visibleName);

        SetCellValue(_fixture, settingsSheet, "$B$4", "C:\\Data");
        AddName(_fixture, visibleName, $"'{settingsSheet}'!$B$4");
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
        Assert.Equal($"='{settingsSheet}'!$B$4", listedRange.RefersTo);
        Assert.Equal("C:\\Data", listedRange.Value);
        Assert.Equal("String", listedRange.ValueType);
        Assert.Equal(1, listedRange.CellCount);
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
            Excel.Names? names = null;
            Excel.Name? nameObject = null;
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
            Excel.Worksheet? sheet = null;
            Excel.Names? names = null;
            Excel.Name? nameObject = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
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
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
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
