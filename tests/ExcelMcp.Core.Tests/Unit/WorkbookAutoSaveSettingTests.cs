using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "Workbook")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class WorkbookAutoSaveSettingTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ApplyAutoSaveSetting_WritesAndVerifiesOnlySelectedWorkbook(bool enabled)
    {
        var selectedSetting = !enabled;
        var otherSetting = !enabled;

        WorkbookCommands.ApplyAutoSaveSetting(
            () => selectedSetting,
            value => selectedSetting = value,
            enabled);

        Assert.Equal(enabled, selectedSetting);
        Assert.Equal(!enabled, otherSetting);
    }

    [Fact]
    public void ApplyAutoSaveSetting_NativeSetterDoesNothing_FailsReadBack()
    {
        var failure = Assert.Throws<InvalidOperationException>(() =>
            WorkbookCommands.ApplyAutoSaveSetting(() => false, _ => { }, enabled: true));

        Assert.Contains("did not apply", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ApplyAutoSaveSetting_StatusUnavailable_FailsBeforeWriting()
    {
        var writeCalled = false;

        Assert.Throws<NotSupportedException>(() =>
            WorkbookCommands.ApplyAutoSaveSetting(() => null, _ => writeCalled = true, enabled: true));

        Assert.False(writeCalled);
    }

    [Fact]
    public void ApplyAutoSaveSetting_NativeFailure_PropagatesOriginalException()
    {
        var nativeFailure = new InvalidOperationException("Native setting failed.");

        var thrown = Assert.Throws<InvalidOperationException>(() =>
            WorkbookCommands.ApplyAutoSaveSetting(
                () => false,
                _ => throw nativeFailure,
                enabled: true));

        Assert.Same(nativeFailure, thrown);
    }
}
