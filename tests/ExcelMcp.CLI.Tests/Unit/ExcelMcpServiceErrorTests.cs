using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

/// <summary>
/// Unit tests for ExcelMcpService error handling.
///
/// REGRESSION TESTS for Bug 5 (GitHub #482): Top-level exception catch in ProcessAsync
/// only included ex.Message, losing the exception type. This makes debugging impossible
/// when the same message text is shared by multiple exception types.
/// </summary>
[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "ExcelMcpService")]
[Trait("Speed", "Fast")]
public sealed class ExcelMcpServiceErrorTests
{
    /// <summary>
    /// REGRESSION TEST for Bug 5 (#482): When an unexpected exception escapes
    /// the ProcessAsync routing switch (e.g. NullReferenceException on null Command),
    /// the error message must include the exception type name so the caller can
    /// distinguish different failure modes.
    /// </summary>
    [Fact]
    public async Task ProcessAsync_UnexpectedExceptionEscapesRouter_ErrorMessageIncludesTypeName()
    {
        // Arrange
        using var service = new ExcelMcpService();

        // null Command triggers NullReferenceException in parts = request.Command.Split(...)
        // This exercises the top-level catch (Exception ex) block in ProcessAsync
#pragma warning disable CS8714 // required property set to null intentionally to trigger NRE
        var request = new ServiceRequest { Command = null! };
#pragma warning restore CS8714

        // Act
        var response = await service.ProcessAsync(request);

        // Assert
        Assert.False(response.Success);
        Assert.NotNull(response.ErrorMessage);

        // REGRESSION: Before fix, only ex.Message was returned ("Object reference not set...").
        // After fix, the type name is prepended: "NullReferenceException: Object reference..."
        Assert.Contains("NullReferenceException", response.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Verifies that normal error responses (business logic, not unexpected exceptions)
    /// still work correctly after the Bug 5 fix. The format change should only affect
    /// the top-level unexpected exception handler.
    /// </summary>
    [Fact]
    public async Task ProcessAsync_UnknownCategory_ReturnsNormalErrorWithoutTypeName()
    {
        // Arrange
        using var service = new ExcelMcpService();
        var request = new ServiceRequest { Command = "unknowncategory.someaction" };

        // Act
        var response = await service.ProcessAsync(request);

        // Assert
        Assert.False(response.Success);
        Assert.NotNull(response.ErrorMessage);
        Assert.Contains("Unknown command category", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        // This path returns a normal string, not an exception-caught message,
        // so it should NOT contain an exception type name prefix.
        Assert.DoesNotContain("Exception:", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Verifies that the WithSessionAsync exception handler (the catch at the bottom
    /// of ProcessAsync, covering session-level operations) also includes the type name.
    /// </summary>
    [Fact]
    public async Task ProcessAsync_SessionCommandWithInvalidSessionId_ReturnsUsableError()
    {
        // Arrange
        using var service = new ExcelMcpService();

        // Send a sheet.list command with a session ID that doesn't exist
        var request = new ServiceRequest
        {
            Command = "sheet.list",
            SessionId = "nonexistent-session-id-00000000"
        };

        // Act
        var response = await service.ProcessAsync(request);

        // Assert — should fail gracefully with a descriptive message, not an unhandled exception
        Assert.False(response.Success);
        Assert.NotNull(response.ErrorMessage);
        Assert.NotEmpty(response.ErrorMessage);
    }
}
