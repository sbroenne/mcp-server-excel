using System.Diagnostics;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Collection("Sequential")]
[Trait("Layer", "ComInterop")]
[Trait("Category", "Integration")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ExcelDialogProbeWindowTests
{
    private static readonly WindowProcedure Procedure = DefWindowProc;

    [Fact]
    public void OwnedModalWindow_IsDetectedBeforeAnOccupiedExcelThread()
    {
        var windowClass = new WindowClass
        {
            WindowProcedure = Marshal.GetFunctionPointerForDelegate(Procedure),
            ClassName = "XLMAIN"
        };
        Assert.NotEqual(0, RegisterClass(ref windowClass));
        nint root = 0;
        nint dialog = 0;
        var failures = new List<Exception>();
        try
        {
            root = CreateWindowEx(0x08000000, "XLMAIN", "", 0, 0, 0, 1, 1, 0, 0, 0, 0);
            Assert.NotEqual(nint.Zero, root);
            using var process = Process.GetCurrentProcess();
            var identity = new ExcelProcessIdentity(
                process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());
            Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Read(identity, NullLogger.Instance));

            dialog = CreateWindowEx(0x08000000, "STATIC", "", 0x90000000, -10000, -10000, 1, 1, root, 0, 0, 0);
            Assert.NotEqual(nint.Zero, dialog);
            EnableWindow(root, false);
            Assert.Equal(WorkbookRefreshState.DialogOpen, ExcelDialogProbe.Read(identity, NullLogger.Instance));

            var batch = (ExcelBatch)RuntimeHelpers.GetUninitializedObject(typeof(ExcelBatch));
            SetField(batch, "_excelProcessIdentity", identity);
            SetField(batch, "_logger", NullLogger<ExcelBatch>.Instance);
            SetField(batch, "_executingWorkItem", 1);
            var started = Stopwatch.GetTimestamp();
            Assert.Equal(WorkbookRefreshState.DialogOpen, batch.GetRefreshState());
            Assert.True(Stopwatch.GetElapsedTime(started) < TimeSpan.FromSeconds(2));

            EnableWindow(root, true);
            Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Read(identity, NullLogger.Instance));
            Assert.Equal(WorkbookRefreshState.Busy, batch.GetRefreshState());
        }
        catch (Exception ex)
        {
            failures.Add(ex);
        }
        finally
        {
            if (dialog != 0 && !DestroyWindow(dialog))
                failures.Add(new Win32Exception("Could not destroy the test dialog."));
            if (root != 0 && !DestroyWindow(root))
                failures.Add(new Win32Exception("Could not destroy the test root window."));
            if (!UnregisterClass("XLMAIN", 0))
                failures.Add(new Win32Exception("Could not unregister the test window class."));
        }
        if (failures.Count == 1) ExceptionDispatchInfo.Capture(failures[0]).Throw();
        if (failures.Count > 1) throw new AggregateException("Dialog test or cleanup failed.", failures);
    }

    private static void SetField(ExcelBatch batch, string name, object value)
    {
        var field = typeof(ExcelBatch).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);
        field.SetValue(batch, value);
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WindowClass
    {
        public uint Style;
        public nint WindowProcedure;
        public int ClassExtra;
        public int WindowExtra;
        public nint Instance;
        public nint Icon;
        public nint Cursor;
        public nint Background;
        public string? MenuName;
        public string ClassName;
    }

    private delegate nint WindowProcedure(nint window, uint message, nint wParam, nint lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern ushort RegisterClass(ref WindowClass windowClass);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnregisterClass(string className, nint instance);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint CreateWindowEx(uint extendedStyle, string className, string title,
        uint style, int x, int y, int width, int height, nint parent, nint menu, nint instance, nint parameter);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyWindow(nint window);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnableWindow(nint window, [MarshalAs(UnmanagedType.Bool)] bool enable);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint DefWindowProc(nint window, uint message, nint wParam, nint lParam);
}
