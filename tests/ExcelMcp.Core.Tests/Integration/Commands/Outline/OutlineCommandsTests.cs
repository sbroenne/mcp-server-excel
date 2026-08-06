using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Outline;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Outline;

[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Outline")]
public sealed class OutlineCommandsTests : IClassFixture<WindowTestsFixture>
{
    private readonly OutlineCommands _commands = new();
    private readonly WindowTestsFixture _fixture;

    public OutlineCommandsTests(WindowTestsFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void SetLevel_Rows_IsExactIdempotentNestedAndClearable()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        var first = _commands.SetLevel(batch, "Sheet1", "2:5", 1, OutlineAxis.Row, collapsed: false);
        Assert.True(first.Success, first.ErrorMessage);
        Assert.Equal(1, first.Level);
        Assert.False(first.Collapsed);
        Assert.True(first.Changed);
        Assert.Equal(4, first.UnitCount);

        var repeated = _commands.SetLevel(batch, "Sheet1", "2:5", 1, OutlineAxis.Row, collapsed: false);
        Assert.False(repeated.Changed);
        Assert.Equal(1, repeated.Level);

        var nested = _commands.SetLevel(batch, "Sheet1", "2:5", 2, OutlineAxis.Row, collapsed: true);
        Assert.Equal(2, nested.Level);
        Assert.True(nested.Collapsed);

        var inspected = _commands.GetState(batch, "Sheet1", "2:5", OutlineAxis.Row);
        Assert.Equal(2, inspected.Level);
        Assert.True(inspected.Collapsed);

        var cleared = _commands.SetLevel(batch, "Sheet1", "2:5", 0, OutlineAxis.Row, collapsed: false);
        Assert.Equal(0, cleared.Level);
        Assert.False(cleared.Collapsed);
    }

    [Fact]
    public void SetLevel_Columns_RoundTripsAndRejectsPartialCellRanges()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        var grouped = _commands.SetLevel(batch, "Sheet1", "B:D", 1, OutlineAxis.Column, collapsed: false);
        Assert.Equal("column", grouped.Axis);
        Assert.Equal(1, grouped.Level);
        Assert.Equal(3, grouped.UnitCount);

        Assert.Throws<ArgumentException>(() =>
            _commands.SetLevel(batch, "Sheet1", "A2:D5", 1, OutlineAxis.Row));

        var stillGrouped = _commands.GetState(batch, "Sheet1", "B:D", OutlineAxis.Column);
        Assert.Equal(1, stillGrouped.Level);

        var cleared = _commands.SetLevel(batch, "Sheet1", "B:D", 0, OutlineAxis.Column, collapsed: false);
        Assert.Equal(0, cleared.Level);
    }
}
