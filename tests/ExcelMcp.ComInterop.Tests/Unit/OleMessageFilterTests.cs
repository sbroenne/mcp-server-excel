using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

/// <summary>
/// Unit tests for OleMessageFilter registration and revocation.
/// Tests verify that the message filter can be registered/revoked without errors.
///
/// NOTE: These tests verify the registration mechanism but don't test actual
/// COM retry behavior (that requires Excel and would be OnDemand tests).
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
public class OleMessageFilterTests
{
    [Fact]
    public void Register_OnStaThread_DoesNotThrow()
    {
        // Arrange & Act & Assert
        var thread = new Thread(() =>
        {
            try
            {
                OleMessageFilter.Register();
                OleMessageFilter.Revoke();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Filter registration failed: {ex.Message}", ex);
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }

    [Fact]
    public void RegisterAndRevoke_MultipleTimes_DoesNotThrow()
    {
        // Arrange & Act & Assert
        var thread = new Thread(() =>
        {
            // First registration
            OleMessageFilter.Register();
            OleMessageFilter.Revoke();

            // Second registration (simulates reuse)
            OleMessageFilter.Register();
            OleMessageFilter.Revoke();
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }

    [Fact]
    public void Revoke_WithoutRegister_DoesNotThrow()
    {
        // Revoke without prior Register should not crash
        // Arrange & Act & Assert - Should handle gracefully
        var thread = new Thread(OleMessageFilter.Revoke);

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }

    /// <summary>
    /// REGRESSION TEST: MessagePending behavior depends on operation mode.
    /// - Normal operations: WAITNOPROCESS (1) — queue messages, don't dispatch
    /// - Long operations: WAITDEFPROCESS (2) — dispatch through HandleInComingCall
    ///
    /// For normal operations, WAITNOPROCESS is correct because dispatching causes
    /// re-entrant COM execution that hangs Data Model operations. The ScreenUpdating
    /// guard in ExcelWriteGuard reduces callbacks, and the deadlock case (FormatConditions.Add
    /// with formula cells) is handled by explicitly wrapping those operations with
    /// EnterLongOperation.
    /// </summary>
    [Fact]
    public void MessagePending_NormalOperation_ReturnsWaitNoProcess()
    {
        const int PENDINGMSG_WAITNOPROCESS = 1;

        var returnValue = -1;
        Exception? threadException = null;

        var thread = new Thread(() =>
        {
            try
            {
                OleMessageFilter.Register();

                Assert.True(OleMessageFilter.IsRegistered, "Filter must be registered to have any effect");

                var filterType = typeof(OleMessageFilter);
                var iOleMsgFilterType = filterType.Assembly.GetType(
                    "Sbroenne.ExcelMcp.ComInterop.IOleMessageFilter");
                Assert.NotNull(iOleMsgFilterType);

                var filterInstance = Activator.CreateInstance(filterType);
                Assert.NotNull(filterInstance);
                var method = iOleMsgFilterType.GetMethod("MessagePending");
                Assert.NotNull(method);

                // Normal operation (not in long operation mode)
                returnValue = (int)method.Invoke(filterInstance, [IntPtr.Zero, 1000, 1])!;
                OleMessageFilter.Revoke();
            }
            catch (Exception ex)
            {
                threadException = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (threadException != null) throw new InvalidOperationException($"Thread exception: {threadException.Message}", threadException);

        Assert.Equal(PENDINGMSG_WAITNOPROCESS, returnValue);
    }

    /// <summary>
    /// Verifies that during long operations, MessagePending returns WAITDEFPROCESS (2)
    /// to dispatch callbacks through HandleInComingCall (which rejects with retry).
    /// </summary>
    [Fact]
    public void MessagePending_DuringLongOperation_ReturnsWaitDefProcess()
    {
        const int PENDINGMSG_WAITDEFPROCESS = 2;

        var returnValue = -1;
        Exception? threadException = null;

        var thread = new Thread(() =>
        {
            try
            {
                OleMessageFilter.Register();
                OleMessageFilter.EnterLongOperation();

                var filterType = typeof(OleMessageFilter);
                var iOleMsgFilterType = filterType.Assembly.GetType(
                    "Sbroenne.ExcelMcp.ComInterop.IOleMessageFilter");
                Assert.NotNull(iOleMsgFilterType);

                var filterInstance = Activator.CreateInstance(filterType);
                Assert.NotNull(filterInstance);
                var method = iOleMsgFilterType.GetMethod("MessagePending");
                Assert.NotNull(method);

                returnValue = (int)method.Invoke(filterInstance, [IntPtr.Zero, 1000, 1])!;

                OleMessageFilter.ExitLongOperation();
                OleMessageFilter.Revoke();
            }
            catch (Exception ex)
            {
                threadException = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (threadException != null) throw new InvalidOperationException($"Thread exception: {threadException.Message}", threadException);

        Assert.Equal(PENDINGMSG_WAITDEFPROCESS, returnValue);
    }
}





