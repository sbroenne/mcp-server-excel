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
        // Arrange & Act & Assert - Should handle gracefully
        var thread = new Thread(() =>
        {
            // Revoke without prior Register should not crash
            OleMessageFilter.Revoke();
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
    }
}
