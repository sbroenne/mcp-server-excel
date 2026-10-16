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

    [Theory]
    [InlineData("Native Database Query - Do you approve running this native query?")]
    [InlineData("Native Datenbankabfrage")]
    [InlineData("Requête de base de données native")]
    public void NativeQueryApprovalWindow_IsDetectedWithinOneSecond(string title)
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
            dialog = CreateWindowEx(0x08000000, "STATIC", title, 0x90000000,
                -10000, -10000, 1, 1, root, 0, 0, 0);
            Assert.NotEqual(nint.Zero, dialog);
            EnableWindow(root, false);

            var started = Stopwatch.GetTimestamp();
            Assert.Equal(WorkbookRefreshState.NativeQueryApproval, ExcelDialogProbe.Read(identity, NullLogger.Instance));
            Assert.True(Stopwatch.GetElapsedTime(started) < TimeSpan.FromSeconds(1));
            var batch = (ExcelBatch)RuntimeHelpers.GetUninitializedObject(typeof(ExcelBatch));
            SetField(batch, "_excelProcessIdentity", identity);
            SetField(batch, "_logger", NullLogger<ExcelBatch>.Instance);
            SetField(batch, "_operationTimeout", TimeSpan.FromSeconds(3));
            SetField(batch, "_workbookPath", "dialog-test.xlsx");
            var timeout = batch.CreateOperationTimeoutException(expiredWhileQueued: false);
            Assert.Contains("waiting for approval to run a native database query", timeout.Message, StringComparison.Ordinal);
            Assert.Contains(title, timeout.Message, StringComparison.Ordinal);
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
        nint message = 0;
        nint input = 0;
        var failures = new List<Exception>();
        try
        {
            root = CreateWindowEx(0x08000000, "XLMAIN", "", 0, 0, 0, 1, 1, 0, 0, 0, 0);
            Assert.NotEqual(nint.Zero, root);
            using var process = Process.GetCurrentProcess();
            var identity = new ExcelProcessIdentity(
                process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());
            Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Read(identity, NullLogger.Instance));

            dialog = CreateWindowEx(0x08000000, "STATIC", "private connection details", 0x90000000, -10000, -10000, 1, 1, root, 0, 0, 0);
            Assert.NotEqual(nint.Zero, dialog);
            message = CreateWindowEx(0, "STATIC", "The database is unavailable.", 0x50000000,
                0, 0, 1, 1, dialog, 0, 0, 0);
            Assert.NotEqual(nint.Zero, message);
            input = CreateWindowEx(0, "EDIT", "test-entered-password", 0x50000000,
                0, 0, 1, 1, dialog, 0, 0, 0);
            Assert.NotEqual(nint.Zero, input);
            EnableWindow(root, false);
            Assert.Equal(WorkbookRefreshState.DialogOpen, ExcelDialogProbe.Read(identity, NullLogger.Instance));

            var batch = (ExcelBatch)RuntimeHelpers.GetUninitializedObject(typeof(ExcelBatch));
            SetField(batch, "_excelProcessIdentity", identity);
            SetField(batch, "_logger", NullLogger<ExcelBatch>.Instance);
            SetField(batch, "_executingWorkItem", 1);
            SetField(batch, "_operationTimeout", TimeSpan.FromSeconds(3));
            SetField(batch, "_workbookPath", "dialog-test.xlsx");
            var started = Stopwatch.GetTimestamp();
            Assert.Equal(WorkbookRefreshState.DialogOpen, batch.GetRefreshState());
            Assert.True(Stopwatch.GetElapsedTime(started) < TimeSpan.FromSeconds(2));

            var timeout = batch.CreateOperationTimeoutException(expiredWhileQueued: false);
            Assert.Contains("had a dialog open", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("timeout recovery may close Excel", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("does not itself cancel", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("private connection details", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("The database is unavailable.", timeout.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("test-entered-password", timeout.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("increasing timeoutSeconds", timeout.Message, StringComparison.Ordinal);
            Assert.Equal(WorkbookRefreshState.DialogOpen, ExcelDialogProbe.Read(identity, NullLogger.Instance));
            var queued = batch.CreateOperationTimeoutException(expiredWhileQueued: true);
            Assert.Contains("was not executed", queued.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("dialog open", queued.Message, StringComparison.Ordinal);

            Assert.True(SetWindowText(message, new string('x', 3000)));
            Assert.Equal(WorkbookRefreshState.DialogOpen,
                ExcelDialogProbe.ReadWithText(identity, NullLogger.Instance, out var truncated));
            Assert.Contains(new string('x', 2047), truncated, StringComparison.Ordinal);
            Assert.DoesNotContain(new string('x', 2048), truncated, StringComparison.Ordinal);
            Assert.Contains("was truncated", truncated, StringComparison.Ordinal);

            var blockedRead = Stopwatch.StartNew();
            string? blockedText = null;
            var blockedState = WorkbookRefreshState.Unknown;
            var reader = new Thread(() =>
                blockedState = ExcelDialogProbe.ReadWithText(identity, NullLogger.Instance, out blockedText));
            reader.Start();
            Assert.True(reader.Join(TimeSpan.FromSeconds(2)));
            Assert.Equal(WorkbookRefreshState.DialogOpen, blockedState);
            Assert.Contains("Some dialog text could not be read", blockedText, StringComparison.Ordinal);
            Assert.True(blockedRead.Elapsed < TimeSpan.FromSeconds(2));

            EnableWindow(root, true);
            Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Read(identity, NullLogger.Instance));
            Assert.Equal(WorkbookRefreshState.Busy, batch.GetRefreshState());
            Assert.Contains("increasing timeoutSeconds",
                batch.CreateOperationTimeoutException(expiredWhileQueued: false).Message, StringComparison.Ordinal);
            SetField(batch, "_excelProcessIdentity",
                new ExcelProcessIdentity(process.Id, identity.StartedAtUtcFileTime - 1));
            var unknown = batch.CreateOperationTimeoutException(expiredWhileQueued: false);
            Assert.DoesNotContain("dialog open", unknown.Message, StringComparison.Ordinal);
            Assert.Contains("increasing timeoutSeconds", unknown.Message, StringComparison.Ordinal);
        }
        catch (Exception ex)
        {
            failures.Add(ex);
        }
        finally
        {
            if (input != 0 && !DestroyWindow(input))
                failures.Add(new Win32Exception("Could not destroy the test input."));
            if (message != 0 && !DestroyWindow(message))
                failures.Add(new Win32Exception("Could not destroy the test message."));
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
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowText(nint window, string text);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint DefWindowProc(nint window, uint message, nint wParam, nint lParam);
}
