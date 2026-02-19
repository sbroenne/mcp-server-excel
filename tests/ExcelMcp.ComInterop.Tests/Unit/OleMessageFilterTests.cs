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
    /// REGRESSION TEST for the STA deadlock bug (Feb 2026):
    /// MessagePending MUST return PENDINGMSG_WAITDEFPROCESS (1), NOT PENDINGMSG_WAITNOPROCESS (2).
    ///
    /// Returning 2 (WAITNOPROCESS) blocks ALL inbound COM message processing while an outgoing
    /// call is in progress. When Excel fires a re-entrant callback (e.g., Calculate, SheetChange)
    /// during FormatConditions.Add(), the callback is queued but WAITNOPROCESS prevents it from
    /// being dispatched. Excel waits for the callback → STA thread waits for Excel → deadlock.
    ///
    /// Returning 1 (WAITDEFPROCESS) allows COM to process the pending inbound call, letting
    /// Excel's callback complete so FormatConditions.Add() can return normally.
    /// </summary>
    [Fact]
    public void MessagePending_ReturnValue_MustBe_WaitDefProcess()
    {
        // The IOleMessageFilter interface is internal, so we verify the constant value via
        // reflection on the compiled method body — simpler: we verify by checking the
        // registered filter doesn't use the blocking value (2).
        //
        // We instantiate the filter and call MessagePending via the interface.
        // Use reflection to access the internal interface implementation.
        const int PENDINGMSG_WAITDEFPROCESS = 1;
        const int PENDINGMSG_WAITNOPROCESS = 2;

        var returnValue = -1;
        Exception? threadException = null;

        var thread = new Thread(() =>
        {
            try
            {
                OleMessageFilter.Register();

                // The filter implements IOleMessageFilter which is internal.
                // We can verify via the public static IsRegistered and the logical behavior:
                // After Register(), the filter IS the active message filter for this thread.
                //
                // Verify that the filter is registered (prerequisite for the bug to manifest).
                Assert.True(OleMessageFilter.IsRegistered, "Filter must be registered to have any effect");

                // Use reflection to invoke MessagePending on the filter instance.
                // The filter class is internal, but we can get to it via the assembly.
                var filterType = typeof(OleMessageFilter);
                var iOleMsgFilterType = filterType.Assembly.GetType(
                    "Sbroenne.ExcelMcp.ComInterop.IOleMessageFilter");
                Assert.NotNull(iOleMsgFilterType);

                // Create a filter instance and call MessagePending
                var filterInstance = Activator.CreateInstance(filterType);
                Assert.NotNull(filterInstance);
                var method = iOleMsgFilterType.GetMethod("MessagePending");
                Assert.NotNull(method);

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

        // REGRESSION: If this returns 2 (WAITNOPROCESS), conditional formatting on cells
        // with formulas will deadlock because Excel's Calculate/SheetChange callbacks
        // can't be delivered while the STA thread waits for FormatConditions.Add().
        Assert.NotEqual(PENDINGMSG_WAITNOPROCESS, returnValue);
        Assert.Equal(PENDINGMSG_WAITDEFPROCESS, returnValue);
    }
}





