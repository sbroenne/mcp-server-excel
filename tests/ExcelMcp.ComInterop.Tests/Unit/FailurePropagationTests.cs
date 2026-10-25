using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop.Formatting;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ErrorHandling")]
[Trait("RequiresExcel", "false")]
[System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
    Justification = "Synthetic COM failures prove the exact exception crosses internal boundaries without calling desktop Excel.")]
public sealed class FailurePropagationTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Formatting_CallerCancellation_Propagates(bool dax)
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => dax
            ? DaxFormatter.FormatAsync("1 + 1", cancellation.Token)
            : MCodeFormatter.FormatAsync("let a = 1 in a", cancellation.Token));
    }

    [Fact]
    public void NumberFormatLocale_UnexpectedComFailure_PreservesOriginalError()
    {
        var error = new COMException("Cannot read locale", unchecked((int)0x80070005));
        var application = DispatchProxy.Create<Excel.Application, ApplicationProxy>();
        Assert.Equal("General", (object?)application.International[Excel.XlApplicationInternational.xlGeneralFormatName]);
        ((ApplicationProxy)(object)application).Error = error;

        Assert.Same(error, Assert.Throws<COMException>(() => new NumberFormatTranslator(application)));
    }

    [Fact]
    public void RefreshProbe_UnexpectedComFailure_PreservesOriginalError()
    {
        var error = new COMException("Cannot read readiness", unchecked((int)0x80070005));
        var application = DispatchProxy.Create<Excel.Application, ApplicationProxy>();
        Assert.True(application.Ready);
        ((ApplicationProxy)(object)application).Error = error;

        Assert.Same(error, Assert.Throws<COMException>(() =>
            WorkbookRefreshProbe.Read(application, null!)));
    }

    [Fact]
    public void WriteGuard_FailedCapture_PropagatesAndDoesNotLeakNesting()
    {
        var error = new COMException("Cannot read ScreenUpdating", unchecked((int)0x80070005));
        var application = DispatchProxy.Create<Excel.Application, ApplicationProxy>();
        var proxy = (ApplicationProxy)(object)application;
        proxy.Error = error;

        Assert.Same(error, Assert.Throws<COMException>(() => new ExcelWriteGuard(application)));
        proxy.Error = null;
        proxy.Workbooks = DispatchProxy.Create<Excel.Workbooks, EmptyWorkbooksProxy>();
        Assert.Equal(0, proxy.Workbooks.Count);
        using (var guard = new ExcelWriteGuard(application))
            Assert.False(proxy.ScreenUpdating);
        Assert.True(proxy.ScreenUpdating);
    }

    public class ApplicationProxy : DispatchProxy
    {
        public Exception? Error { get; set; }
        public bool ScreenUpdating { get; private set; } = true;
        public Excel.Workbooks? Workbooks { get; set; }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (Error != null)
                throw Error;
            return targetMethod?.Name switch
            {
                "get_ScreenUpdating" => ScreenUpdating,
                "get_Ready" => true,
                "get_International" => "General",
                "get_Workbooks" => Workbooks,
                "set_ScreenUpdating" => SetScreenUpdating((bool)args![0]!),
                _ => throw new NotSupportedException(targetMethod?.Name)
            };
        }

        private object? SetScreenUpdating(bool value)
        {
            ScreenUpdating = value;
            return null;
        }
    }

    public class EmptyWorkbooksProxy : DispatchProxy
    {
        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args) =>
            targetMethod?.Name == "get_Count" ? 0 : throw new NotSupportedException(targetMethod?.Name);
    }
}
