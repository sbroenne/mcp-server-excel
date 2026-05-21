using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.NamedRange;

/// <summary>
/// Regression tests for named range list behavior on workbooks with hidden/internal names and large ranges.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public sealed class NamedRangeCommandsListRegressionTests : IClassFixture<TempDirectoryFixture>
{
    private readonly NamedRangeCommands _commands = new();
    private readonly TempDirectoryFixture _fixture;

    public NamedRangeCommandsListRegressionTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void List_HiddenUserDefinedName_DoesNotExposeHiddenNameByDefault()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        var name = NamedRangeTestsFixture.GetUniqueNamedRangeName();

        AddHiddenUserDefinedName(batch, name);

        var result = _commands.List(batch);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.DoesNotContain(result.NamedRanges, namedRange => namedRange.Name == name);
    }

    [Fact]
    public void List_LargeNamedRange_OmitsValuePreview()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        var name = NamedRangeTestsFixture.GetUniqueNamedRangeName();

        AddNamedRange(batch, name, "Sheet1!$A$1:$A$10001");

        var result = _commands.List(batch);

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
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        var visibleName = NamedRangeTestsFixture.GetUniqueNamedRangeName();

        AddWorksheet(batch, "Users");
        AddWorksheet(batch, "Notifications");
        SetCellValue(batch, "Sheet1", "$B$4", "C:\\Data");
        AddNamedRange(batch, visibleName, "Sheet1!$B$4");
        AddHiddenSheetScopedName(batch, "Users", "ExternalData_1", "Users!$A$6:$AH$19132");
        AddHiddenSheetScopedName(batch, "Notifications", "ExternalData_1", "Notifications!$A$6:$P$28365");
        batch.Save();

        var result = _commands.List(batch);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        var listedRange = Assert.Single(result.NamedRanges);
        Assert.Equal(visibleName, listedRange.Name);
        Assert.DoesNotContain(result.NamedRanges, namedRange => namedRange.Name.Contains("ExternalData_1", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void List_JapaneseSheetsWithMultipleLargeHiddenExternalDataNames_ReturnsVisibleUserName()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);
        var visibleName = NamedRangeTestsFixture.GetUniqueNamedRangeName();

        AddWorksheet(batch, "PQ_設定");
        AddWorksheet(batch, "ユーザーテーブル_users");
        AddWorksheet(batch, "通知テーブル_notifications");
        SetCellValue(batch, "PQ_設定", "$B$4", "C:\\Data");
        AddNamedRange(batch, visibleName, "PQ_設定!$B$4");
        AddHiddenSheetScopedName(batch, "ユーザーテーブル_users", "ExternalData_1", "ユーザーテーブル_users!$A$6:$AH$19132");
        AddHiddenSheetScopedName(batch, "通知テーブル_notifications", "ExternalData_1", "通知テーブル_notifications!$A$6:$P$28365");
        batch.Save();

        var result = _commands.List(batch);

        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        var listedRange = Assert.Single(result.NamedRanges);
        Assert.Equal(visibleName, listedRange.Name);
        Assert.Contains("PQ_設定", listedRange.RefersTo, StringComparison.Ordinal);
        Assert.Equal("C:\\Data", listedRange.Value);
        Assert.DoesNotContain(result.NamedRanges, namedRange => namedRange.Name.Contains("ExternalData_1", StringComparison.OrdinalIgnoreCase));
    }

    private static void AddHiddenUserDefinedName(IExcelBatch batch, string name)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? names = null;
            dynamic? nameObj = null;
            try
            {
                names = ctx.Book.Names;
                nameObj = names.Add(name, "=Sheet1!$B$4");
                nameObj.Visible = false;
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
                ComUtilities.Release(ref names);
            }
        });
    }

    private static void AddNamedRange(IExcelBatch batch, string name, string reference)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? names = null;
            dynamic? nameObj = null;
            try
            {
                names = ctx.Book.Names;
                nameObj = names.Add(name, $"={reference.TrimStart('=')}");
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
                ComUtilities.Release(ref names);
            }
        });
    }

    private static void AddWorksheet(IExcelBatch batch, string sheetName)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? sheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = sheets.Add();
                sheet.Name = sheetName;
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    private static void SetCellValue(IExcelBatch batch, string sheetName, string rangeAddress, string value)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                range = sheet.Range[rangeAddress];
                range.Value2 = value;
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void AddHiddenSheetScopedName(IExcelBatch batch, string sheetName, string name, string reference)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? names = null;
            dynamic? nameObj = null;
            try
            {
                names = ctx.Book.Names;
                nameObj = names.Add($"{sheetName}!{name}", $"={reference.TrimStart('=')}");
                nameObj.Visible = false;
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref nameObj);
                ComUtilities.Release(ref names);
            }
        });
    }
}
