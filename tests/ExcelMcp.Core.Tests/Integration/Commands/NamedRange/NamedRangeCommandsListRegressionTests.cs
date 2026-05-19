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
}
