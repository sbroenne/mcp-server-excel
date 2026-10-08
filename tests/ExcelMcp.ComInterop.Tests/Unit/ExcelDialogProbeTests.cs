using System.Diagnostics;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelDialogProbeTests
{
    [Theory]
    [InlineData(42)]
    [InlineData(99)]
    public void VisibleModalOwnedByDisabledExcel_DetectedAcrossProcesses(int dialogProcessId)
    {
        ExcelDialogProbe.WindowSnapshot[] windows =
        [
            new(1, 0, 42, true, false, "XLMAIN"),
            new(2, 1, dialogProcessId, true, true, "ProxyModalWindow")
        ];
        Assert.Equal(WorkbookRefreshState.DialogOpen, ExcelDialogProbe.Classify(42, windows));
    }

    [Fact]
    public void DialogOwnerChain_CanIncludeAnotherProcessAndHiddenOwner()
    {
        ExcelDialogProbe.WindowSnapshot[] windows =
        [
            new(1, 0, 42, false, false, "XLMAIN"),
            new(2, 1, 99, false, false, "ProxyModalWindow"),
            new(3, 2, 100, true, true, "#32770")
        ];
        Assert.Equal(WorkbookRefreshState.DialogOpen, ExcelDialogProbe.Classify(42, windows));
    }

    [Theory]
    [InlineData(true, true, true)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    public void ModelessHiddenOrDisabledPopup_DoesNotEstablishADialog(
        bool excelEnabled, bool popupVisible, bool popupEnabled)
    {
        ExcelDialogProbe.WindowSnapshot[] windows =
        [
            new(1, 0, 42, true, excelEnabled, "XLMAIN"),
            new(2, 1, 42, popupVisible, popupEnabled, "#32770")
        ];
        Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Classify(42, windows));
    }

    [Fact]
    public void OtherExcelInstanceDialog_DoesNotBlockIntendedInstance()
    {
        ExcelDialogProbe.WindowSnapshot[] windows =
        [
            new(1, 0, 42, true, false, "XLMAIN"),
            new(2, 0, 99, true, false, "XLMAIN"),
            new(3, 2, 99, true, true, "#32770")
        ];
        Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Classify(42, windows));
        Assert.Equal(WorkbookRefreshState.DialogOpen, ExcelDialogProbe.Classify(99, windows));
    }

    [Fact]
    public void DisabledExcelWithoutAnOwnedPopup_DoesNotInventUserInput()
    {
        ExcelDialogProbe.WindowSnapshot[] windows = [new(1, 0, 42, true, false, "XLMAIN")];
        Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Classify(42, windows));
    }

    [Fact]
    public void CyclicOrMissingOwners_DoNotEstablishADialog()
    {
        ExcelDialogProbe.WindowSnapshot[] windows =
        [
            new(1, 0, 42, true, false, "XLMAIN"),
            new(2, 3, 99, true, true, "#32770"),
            new(3, 2, 99, true, true, "#32770"),
            new(4, 999, 99, true, true, "#32770")
        ];
        Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Classify(42, windows));
    }

    [Fact]
    public void MissingExcelWindow_IsUnknown()
    {
        Assert.Equal(WorkbookRefreshState.Unknown, ExcelDialogProbe.Classify(42, []));
    }

    [Fact]
    public void MissingOrReusedProcessIdentity_IsUnknown()
    {
        Assert.Equal(WorkbookRefreshState.Unknown, ExcelDialogProbe.Read(null, NullLogger.Instance));
        using var process = Process.GetCurrentProcess();
        var reused = new ExcelProcessIdentity(
            process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc() + 1);
        Assert.Equal(WorkbookRefreshState.Unknown, ExcelDialogProbe.Read(reused, NullLogger.Instance));
    }

    [Theory]
    [InlineData("save")]
    [InlineData("close")]
    [InlineData("save as")]
    public void DialogOpen_ExplainsPromptAndPreservesUnsavedState(string operation)
    {
        var exception = Assert.Throws<ExcelBusyException>(() =>
            ExcelBusyException.ThrowIfNotReady(WorkbookRefreshState.DialogOpen, operation));
        Assert.Contains($"Cannot {operation}", exception.Message, StringComparison.Ordinal);
        Assert.Contains("Check the Excel window", exception.Message, StringComparison.Ordinal);
        Assert.Contains("not been saved or discarded", exception.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("sign-in required", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
