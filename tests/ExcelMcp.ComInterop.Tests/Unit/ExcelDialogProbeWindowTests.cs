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
    public async Task Show_FromHiddenWindow_ConfirmsRequestedVisibility()
    {
        var windowClass = new WindowClass
        {
            WindowProcedure = Marshal.GetFunctionPointerForDelegate(Procedure),
            ClassName = "XLMAIN"
        };
        Assert.NotEqual(0, RegisterClass(ref windowClass));
        nint root = 0;
        var failures = new List<Exception>();
        try
        {
            root = CreateWindowEx(0x08000000, "XLMAIN", "", 0x80000000,
                -10000, -10000, 1, 1, 0, 0, 0, 0);
            Assert.NotEqual(nint.Zero, root);
            Assert.False(IsWindowVisible(root));
            using var process = Process.GetCurrentProcess();
            var identity = new ExcelProcessIdentity(
                process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());

            var visibility = Task.Run(() =>
                ExcelWindowVisibility.Set(identity, true, NullLogger.Instance));
            var timer = Stopwatch.StartNew();
            while (!visibility.IsCompleted && timer.Elapsed < TimeSpan.FromSeconds(5))
            {
                if (PeekMessage(out var message, 0, 0, 0, 1))
                    _ = DispatchMessage(ref message);
                else
                    Thread.Sleep(1);
            }

            var result = await visibility.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.True(result.IsVisible);
            Assert.True(IsWindowVisible(root));
        }
        catch (Exception ex)
        {
            failures.Add(ex);
        }
        finally
        {
            if (root != 0 && !DestroyWindow(root))
                failures.Add(new Win32Exception("Could not destroy the test root window."));
            if (!UnregisterClass("XLMAIN", 0))
                failures.Add(new Win32Exception("Could not unregister the test window class."));
        }
        if (failures.Count == 1) ExceptionDispatchInfo.Capture(failures[0]).Throw();
        if (failures.Count > 1) throw new AggregateException("Visibility test or cleanup failed.", failures);
    }

    [Fact]
    public void Show_WhenMinimizedWindowCannotProcessRestore_ReportsUnconfirmedVisibility()
    {
        var windowClass = new WindowClass
        {
            WindowProcedure = Marshal.GetFunctionPointerForDelegate(Procedure),
            ClassName = "XLMAIN"
        };
        Assert.NotEqual(0, RegisterClass(ref windowClass));
        nint root = 0;
        var failures = new List<Exception>();
        try
        {
            root = CreateWindowEx(0x08000000, "XLMAIN", "", 0xB0000000,
                -10000, -10000, 1, 1, 0, 0, 0, 0);
            Assert.NotEqual(nint.Zero, root);
            Assert.True(IsWindowVisible(root));
            Assert.True(IsIconic(root));
            using var process = Process.GetCurrentProcess();
            var identity = new ExcelProcessIdentity(
                process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());

            // Do not pump the owning thread: Excel may likewise be unable to process a restore.
            var visibility = Task.Run(() => ExcelWindowVisibility.Set(identity, true, NullLogger.Instance));
            var failure = Record.Exception(() =>
                visibility.WaitAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult());
            Assert.Contains("did not confirm", Assert.IsType<InvalidOperationException>(failure).Message,
                StringComparison.Ordinal);
            Assert.True(IsWindowVisible(root));
            Assert.True(IsIconic(root));
        }
        catch (Exception ex)
        {
            failures.Add(ex);
        }
        finally
        {
            if (root != 0 && !DestroyWindow(root))
                failures.Add(new Win32Exception("Could not destroy the test root window."));
            if (!UnregisterClass("XLMAIN", 0))
                failures.Add(new Win32Exception("Could not unregister the test window class."));
        }
        if (failures.Count == 1) ExceptionDispatchInfo.Capture(failures[0]).Throw();
        if (failures.Count > 1) throw new AggregateException("Visibility test or cleanup failed.", failures);
    }

    [Fact]
    public void Hide_WhenExcelReplacesItsWindow_ReportsChangedOwnershipInsteadOfSuccess()
    {
        nint root = 0;
        var replaced = false;
        using var replacementFinished = new ManualResetEventSlim();
        var failures = new List<Exception>();
        WindowProcedure replacingProcedure = (window, message, wParam, lParam) =>
        {
            if (message == 0x46 && window == root && !replaced &&
                (Marshal.ReadInt32(lParam, 2 * IntPtr.Size + 4 * sizeof(int)) & 0x80) != 0)
            {
                replaced = true;
                if (!DestroyWindow(root))
                    failures.Add(new Win32Exception("Could not replace the test root window."));
                root = CreateWindowEx(0x08000000, "XLMAIN", "", 0x90000000,
                    -10000, -10000, 1, 1, 0, 0, 0, 0);
                replacementFinished.Set();
                return 0;
            }
            return DefWindowProc(window, message, wParam, lParam);
        };
        var windowClass = new WindowClass
        {
            WindowProcedure = Marshal.GetFunctionPointerForDelegate(replacingProcedure),
            ClassName = "XLMAIN"
        };
        Assert.NotEqual(0, RegisterClass(ref windowClass));
        try
        {
            root = CreateWindowEx(0x08000000, "XLMAIN", "", 0x90000000,
                -10000, -10000, 1, 1, 0, 0, 0, 0);
            Assert.NotEqual(nint.Zero, root);
            using var process = Process.GetCurrentProcess();
            var identity = new ExcelProcessIdentity(
                process.Id, process.StartTime.ToUniversalTime().ToFileTimeUtc());
            ExcelWindowVisibility.BeforeVisibilityConfirmationHookForTests = () =>
                Assert.True(replacementFinished.Wait(TimeSpan.FromSeconds(5)));
            var visibility = Task.Run(() => ExcelWindowVisibility.Set(identity, false, NullLogger.Instance));
            var timer = Stopwatch.StartNew();
            while (!visibility.IsCompleted && timer.Elapsed < TimeSpan.FromSeconds(5))
            {
                if (PeekMessage(out var message, 0, 0, 0, 1))
                    _ = DispatchMessage(ref message);
                else
                    Thread.Sleep(1);
            }
            var failure = Record.Exception(() => visibility.WaitAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult());
            Assert.True(replaced);
            Assert.NotEqual(nint.Zero, root);
            Assert.True(IsWindowVisible(root));
            Assert.Contains("changed", Assert.IsType<InvalidOperationException>(failure).Message, StringComparison.Ordinal);
        }
        catch (Exception ex)
        {
            failures.Add(ex);
        }
        finally
        {
            ExcelWindowVisibility.BeforeVisibilityConfirmationHookForTests = null;
            if (root != 0 && !DestroyWindow(root))
                failures.Add(new Win32Exception("Could not destroy the test root window."));
            if (!UnregisterClass("XLMAIN", 0))
                failures.Add(new Win32Exception("Could not unregister the test window class."));
            GC.KeepAlive(replacingProcedure);
        }
        if (failures.Count == 1) ExceptionDispatchInfo.Capture(failures[0]).Throw();
        if (failures.Count > 1) throw new AggregateException("Visibility test or cleanup failed.", failures);
    }

    [Theory]
    [InlineData("Native Database Query - Do you approve running this native query?")]
    [InlineData("Native Datenbankabfrage")]
    [InlineData("Requête de base de données native")]
    public void DialogWindow_ReportsReadableTitleWithoutClassifyingItsType(string title)
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
            Assert.Equal(WorkbookRefreshState.DialogOpen, ExcelDialogProbe.Read(identity, NullLogger.Instance));
            Assert.True(Stopwatch.GetElapsedTime(started) < TimeSpan.FromSeconds(1));
            var batch = (ExcelBatch)RuntimeHelpers.GetUninitializedObject(typeof(ExcelBatch));
            SetField(batch, "_excelProcessIdentity", identity);
            SetField(batch, "_logger", NullLogger<ExcelBatch>.Instance);
            SetField(batch, "_operationTimeout", TimeSpan.FromSeconds(3));
            SetField(batch, "_workbookPath", "dialog-test.xlsx");
            var timeout = batch.CreateOperationTimeoutException(expiredWhileQueued: false);
            Assert.Contains("had a dialog open", timeout.Message, StringComparison.Ordinal);
            Assert.Contains(title, timeout.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("waiting for approval", timeout.Message, StringComparison.Ordinal);
            TimeoutException? blockedTimeout = null;
            var blockedRead = new Thread(() => blockedTimeout = batch.CreateOperationTimeoutException(expiredWhileQueued: false));
            blockedRead.Start();
            Assert.True(blockedRead.Join(TimeSpan.FromSeconds(2)));
            Assert.Contains("could not be read", Assert.IsType<TimeoutException>(blockedTimeout).Message, StringComparison.Ordinal);
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
    public async Task OwnedModalWindow_IsDetectedBeforeAnOccupiedExcelThread()
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
            Assert.Equal("No blocking Excel dialog was detected.",
                ExcelDialogProbe.ReadDetails(identity, NullLogger.Instance));

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
            SetField(batch, "_operationTimeout", TimeSpan.FromSeconds(3));
            SetField(batch, "_workbookPath", "dialog-test.xlsx");
            SetField(batch, "_windowControlLock", new Lock());
            var started = Stopwatch.GetTimestamp();
            SetField(batch, "_startupCompletion", Task.CompletedTask);
            SetField(batch, "_statusLock", new Lock());
            SetField(batch, "_operationState", "running");
            var detailRead = Task.Run(() => batch.DialogDetails);
            while (!detailRead.IsCompleted && Stopwatch.GetElapsedTime(started) < TimeSpan.FromSeconds(5))
            {
                if (PeekMessage(out var pendingMessage, 0, 0, 0, 1))
                    _ = DispatchMessage(ref pendingMessage);
                else
                    Thread.Sleep(1);
            }
            var details = await detailRead.WaitAsync(TimeSpan.FromSeconds(1));
            Assert.Contains("private connection details", details, StringComparison.Ordinal);
            Assert.Contains("The database is unavailable.", details, StringComparison.Ordinal);
            Assert.DoesNotContain("test-entered-password", details, StringComparison.Ordinal);
            Assert.True(IsWindowVisible(root), batch.DialogVisibilityError);
            Assert.Equal(WorkbookRefreshState.DialogOpen, batch.GetRefreshState());
            Assert.True(Stopwatch.GetElapsedTime(started) < TimeSpan.FromSeconds(2));

            var timeout = batch.CreateOperationTimeoutException(expiredWhileQueued: false);
            Assert.Contains("had a dialog open", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("workbook remains open", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("No changes were saved or discarded by timeout recovery", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("does not itself cancel", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("private connection details", timeout.Message, StringComparison.Ordinal);
            Assert.Contains("The database is unavailable.", timeout.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("test-entered-password", timeout.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("increasing timeoutSeconds", timeout.Message, StringComparison.Ordinal);
            Assert.Equal(WorkbookRefreshState.DialogOpen, ExcelDialogProbe.Read(identity, NullLogger.Instance));
            var queued = batch.CreateOperationTimeoutException(expiredWhileQueued: true);
            Assert.Contains("was not executed", queued.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("dialog open", queued.Message, StringComparison.Ordinal);

            for (int i = 0; i < 5; i++)
                Assert.NotEqual(nint.Zero, CreateWindowEx(0, "STATIC", new string('x', 2000), 0x50000000,
                    0, 0, 1, 1, dialog, 0, 0, 0));
            var truncated = ExcelDialogProbe.ReadDetails(identity, NullLogger.Instance);
            Assert.Contains("truncated", truncated, StringComparison.Ordinal);
            var captured = truncated["Dialog text: ".Length..].Split("\n[Some dialog text", 2)[0];
            Assert.Equal(8192, captured.Length);

            var blockedRead = Stopwatch.StartNew();
            var blockedState = WorkbookRefreshState.Unknown;
            var reader = new Thread(() =>
                blockedState = ExcelDialogProbe.Read(identity, NullLogger.Instance));
            reader.Start();
            Assert.True(reader.Join(TimeSpan.FromSeconds(2)));
            Assert.Equal(WorkbookRefreshState.DialogOpen, blockedState);
            Assert.True(blockedRead.Elapsed < TimeSpan.FromSeconds(2));

            EnableWindow(root, true);
            Assert.Equal(WorkbookRefreshState.Ready, ExcelDialogProbe.Read(identity, NullLogger.Instance));
            Assert.Equal("No blocking Excel dialog was detected.",
                ExcelDialogProbe.ReadDetails(identity, NullLogger.Instance));
            Assert.Equal(WorkbookRefreshState.Busy, batch.GetRefreshState());
            Assert.Contains("Inspect session status",
                batch.CreateOperationTimeoutException(expiredWhileQueued: false).Message, StringComparison.Ordinal);
            SetField(batch, "_excelProcessIdentity",
                new ExcelProcessIdentity(process.Id, identity.StartedAtUtcFileTime - 1));
            Assert.Contains("could not confirm", batch.DialogDetails, StringComparison.Ordinal);
            var unknown = batch.CreateOperationTimeoutException(expiredWhileQueued: false);
            Assert.DoesNotContain("dialog open", unknown.Message, StringComparison.Ordinal);
            Assert.Contains("Inspect session status", unknown.Message, StringComparison.Ordinal);
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

    private static void SetField<T>(ExcelBatch batch, string name, T value)
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

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMessage
    {
        public nint Window;
        public uint Message;
        public nint WParam;
        public nint LParam;
        public uint Time;
        public int X;
        public int Y;
        public uint Private;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PeekMessage(out NativeMessage message, nint window, uint min, uint max, uint remove);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint DispatchMessage(ref NativeMessage message);

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

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(nint window);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsIconic(nint window);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint DefWindowProc(nint window, uint message, nint wParam, nint lParam);
}
