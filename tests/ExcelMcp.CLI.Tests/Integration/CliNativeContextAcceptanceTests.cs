using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "Protection")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeProtectionAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task NativeCellAndSheetProtection_PreserveUnselectedGapAndPermissions()
    {
        await RangeAsync("rangelink", "set-cell-protection", "S1,S3", "--locked", "false", "--formula-hidden", "true");
        var cells = await RangeAsync("rangelink", "get-cell-protection", "S1:S3");
        Assert.Equal(3, Items(cells, "cells").Length);
        Assert.False(Bool(cells, "cells.0.locked"));
        Assert.True(Bool(cells, "cells.0.formulaHidden"));
        Assert.True(Bool(cells, "cells.1.locked"));
        Assert.False(Bool(cells, "cells.1.formulaHidden"));
        await WithCleanupAsync(async () =>
        {
            await CommandAsync("worksheetstyle", "set-protection", "--sheet-name", "Data", "--is-protected", "true",
                "--options", """{"allowFormattingRows":true}""");
            var permissions = await CommandAsync("worksheetstyle", "get-protection", "--sheet-name", "Data");
            Assert.True(Bool(permissions, "protectContents"));
            Assert.True(Bool(permissions, "permissions.allowFormattingRows"));
            Assert.False(Bool(permissions, "permissions.allowSorting"));
        }, async () => { await CommandAsync("worksheetstyle", "set-protection", "--sheet-name", "Data", "--is-protected", "false"); });
    }
}

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "CalculationMode")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeCalculationAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task NativeSettingsAndRebuild_RestoreOriginalCalculationMode()
    {
        var original = await CommandAsync("calculationmode", "get-settings");
        Assert.Equal("application", Text(original, "settingsScope"));
        Assert.True(Number(original, "maximumIterations") > 0);
        Assert.False(Bool(original, "precisionAsDisplayed"));
        var mode = Text(original, "mode");
        await WithCleanupAsync(async () =>
        {
            Assert.Equal("manual", Text(await CommandAsync("calculationmode", "set-settings", "--mode", "manual"), "mode"));
            await CommandAsync("calculationmode", "calculate", "--scope", "application", "--kind", "rebuild");
        }, async () =>
        {
            Assert.Equal(mode, Text(await CommandAsync("calculationmode", "set-settings", "--mode", mode), "mode"));
        });
        Assert.False(Bool(await CommandAsync("calculationmode", "set-precision", "--precision-as-displayed", "false"), "precisionAsDisplayed"));
    }
}

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "WorkbookTheme")]
[Trait("Feature", "Window")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeContextAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task NativeOwnedWindowAndTheme_ExposeCompleteDefinitions()
    {
        var context = await CommandAsync("window", "get-context");
        Assert.Equal("available", Text(context, "availability"));
        Assert.NotEmpty(Items(context, "windows"));
        var theme = await CommandAsync("workbook", "get-theme");
        Assert.Equal(12, Items(theme, "colors").Length);
        Assert.Equal(3, Items(theme, "majorFonts").Length);
        Assert.Equal(3, Items(theme, "minorFonts").Length);
    }
}
