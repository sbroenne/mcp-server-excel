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
    [InlineData("Native Database Query - Do you approve running this native query?")]
    [InlineData("Native Datenbankabfrage")]
    [InlineData("Requête de base de données native")]
    public void OwnedDialog_IsGenericRegardlessOfTitle(string title)
    {
        ExcelDialogProbe.WindowSnapshot[] windows =
        [
            new(1, 0, 42, true, false, "XLMAIN"),
            new(2, 1, 42, true, true, "#32770", title)
        ];

        Assert.Equal(WorkbookRefreshState.DialogOpen, ExcelDialogProbe.Classify(42, windows));
    }

    [Fact]
    public void NativeQueryApprovalTitleWithoutOwnedModal_DoesNotBlockExcel()
    {
        ExcelDialogProbe.WindowSnapshot[] windows =
        [
            new(1, 0, 42, true, true, "XLMAIN"),
            new(2, 0, 42, true, true, "#32770", "Native Database Query")
        ];

        Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Classify(42, windows));
    }

    [Fact]
    public void OtherOwnedModalDialog_RemainsGeneric()
    {
        ExcelDialogProbe.WindowSnapshot[] windows =
        [
            new(1, 0, 42, true, false, "XLMAIN"),
            new(2, 1, 42, true, true, "#32770", "Sign in")
        ];

        Assert.Equal(WorkbookRefreshState.DialogOpen, ExcelDialogProbe.Classify(42, windows));
    }

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
    public void ModelessFindReplaceWindow_DoesNotBlockReadinessOrSaveClose()
    {
        ExcelDialogProbe.WindowSnapshot[] windows =
        [
            new(1, 0, 42, true, true, "XLMAIN"),
            new(2, 1, 42, true, true, "bosa_sdm_XL9", "Find and Replace")
        ];
        Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Classify(42, windows));
        Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Classify(42, windows, out _));
        foreach (var operation in new[] { "save", "close" })
            ExcelBusyException.ThrowIfNotReady(ExcelDialogProbe.Classify(42, windows), operation);
    }

    [Fact]
    public void InputBoxWithEnabledExcelOwner_IsDetectedOnlyForTimeoutDiagnostics()
    {
        ExcelDialogProbe.WindowSnapshot[] windows =
        [
            new(1, 0, 42, true, true, "XLMAIN"),
            new(2, 1, 42, true, true, "bosa_sdm_XL9", "Input")
        ];

        Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Classify(42, windows));
        Assert.Equal(
            WorkbookRefreshState.DialogOpen,
            ExcelDialogProbe.Classify(42, windows, out _, includeEnabledOwnerOfficeDialog: true));
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
            new(3, 2, 99, true, true, "#32770", "Native Database Query")
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

    [Theory]
    [InlineData("Dialog title: Test approval. Dialog text: Review this request.", null)]
    [InlineData("Dialog text is incomplete because one control did not respond.", "Excel could not be revealed.")]
    [InlineData("Excel dialog inspection could not confirm the session's window ownership.", null)]
    [InlineData(null, "Excel could not be revealed.")]
    [InlineData("", null)]
    public void DialogOpen_IncludesAvailableDetailsAndVisibilityFailure(string? details, string? visibilityError)
    {
        var window = new DialogWindowControl(details, visibilityError);
        var exception = Assert.Throws<ExcelBusyException>(() =>
            ExcelBusyException.ThrowIfNotReady(WorkbookRefreshState.DialogOpen, "save", window));

        Assert.Contains(string.IsNullOrWhiteSpace(details) ? "Readable dialog text is unavailable." : details,
            exception.Message, StringComparison.Ordinal);
        if (visibilityError is not null)
            Assert.Contains(visibilityError, exception.Message, StringComparison.Ordinal);
        Assert.Contains("not been saved or discarded", exception.Message, StringComparison.Ordinal);
        Assert.Equal(1, window.DetailReads);
        Assert.Equal(1, window.VisibilityErrorReads);
    }

    [Fact]
    public void OtherStates_DoNotReadDialogDetailsOrVisibilityErrors()
    {
        foreach (var state in Enum.GetValues<WorkbookRefreshState>().Where(state => state != WorkbookRefreshState.DialogOpen))
        {
            var window = new DialogWindowControl("Unrelated dialog text", "Unrelated visibility error");
            if (state == WorkbookRefreshState.Ready)
                ExcelBusyException.ThrowIfNotReady(state, "save", window);
            else
            {
                var exception = Assert.Throws<ExcelBusyException>(() =>
                    ExcelBusyException.ThrowIfNotReady(state, "save", window));
                Assert.DoesNotContain("Unrelated", exception.Message, StringComparison.Ordinal);
            }
            Assert.Equal(0, window.DetailReads);
            Assert.Equal(0, window.VisibilityErrorReads);
        }
    }

    [Fact]
    public void DialogOpen_WithoutWindowControl_ExplainsUnavailableText()
    {
        var exception = Assert.Throws<ExcelBusyException>(() =>
            ExcelBusyException.ThrowIfNotReady(WorkbookRefreshState.DialogOpen, "save"));
        Assert.Contains("Readable dialog text is unavailable.", exception.Message, StringComparison.Ordinal);
    }

    private sealed class DialogWindowControl(string? details, string? visibilityError) : IExcelBatchWindowControl
    {
        internal int DetailReads { get; private set; }
        internal int VisibilityErrorReads { get; private set; }

        public string? DialogDetails
        {
            get
            {
                DetailReads++;
                return details;
            }
        }

        public string? DialogVisibilityError
        {
            get
            {
                VisibilityErrorReads++;
                return visibilityError;
            }
        }

        public ExcelVisibilityResult SetVisibility(bool visible) =>
            throw new InvalidOperationException("The error builder must not change visibility itself.");
    }
}
