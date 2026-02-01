// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.ExcelMcp.Core.Models.Actions;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Integration test that verifies telemetry configuration and sensitive data redaction.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Telemetry")]
public class TelemetryIntegrationTests(ITestOutputHelper output)
{
    [Fact]
    public void TelemetryConfiguration_HasStableUserAndSessionIds()
    {
        output.WriteLine("=== TELEMETRY CONFIGURATION TEST ===\n");

        // Get user and session IDs
        var userId = ExcelMcpTelemetry.UserId;
        var sessionId = ExcelMcpTelemetry.SessionId;

        output.WriteLine($"User ID: {userId}");
        output.WriteLine($"Session ID: {sessionId}");

        // Assert - user ID should be stable (16 hex chars from SHA256)
        Assert.NotNull(userId);
        Assert.Equal(16, userId.Length);
        Assert.True(userId.All(c => char.IsAsciiHexDigitLower(c)), "User ID should be lowercase hex");

        // Assert - session ID should be unique per process (8 hex chars from GUID)
        Assert.NotNull(sessionId);
        Assert.Equal(8, sessionId.Length);
        Assert.True(sessionId.All(c => char.IsAsciiHexDigit(c)), "Session ID should be hex");

        // Verify IDs are consistent within same process
        Assert.Equal(userId, ExcelMcpTelemetry.UserId);
        Assert.Equal(sessionId, ExcelMcpTelemetry.SessionId);
    }

    [Fact]
    public void SensitiveDataRedactor_RedactsFilePaths()
    {
        var input = "Error loading file C:\\Users\\John\\Documents\\secret.xlsx";
        var redacted = SensitiveDataRedactor.RedactSensitiveData(input);

        output.WriteLine($"Input: {input}");
        output.WriteLine($"Redacted: {redacted}");

        Assert.DoesNotContain("C:\\", redacted);
        Assert.Contains("[REDACTED_PATH]", redacted);
    }

    [Fact]
    public void SensitiveDataRedactor_RedactsConnectionStrings()
    {
        var input = "Connection: Server=myserver;Password=secret123;User=admin";
        var redacted = SensitiveDataRedactor.RedactSensitiveData(input);

        output.WriteLine($"Input: {input}");
        output.WriteLine($"Redacted: {redacted}");

        Assert.DoesNotContain("secret123", redacted);
        Assert.Contains("[REDACTED]", redacted);
    }

    [Fact]
    public void SensitiveDataRedactor_RedactsEmailAddresses()
    {
        var input = "Contact john.doe@example.com for support";
        var redacted = SensitiveDataRedactor.RedactSensitiveData(input);

        output.WriteLine($"Input: {input}");
        output.WriteLine($"Redacted: {redacted}");

        Assert.DoesNotContain("john.doe@example.com", redacted);
        Assert.Contains("[REDACTED_EMAIL]", redacted);
    }

    [Fact]
    public void SensitiveDataRedactor_RedactsExceptions()
    {
        var exception = new InvalidOperationException("Failed to read C:\\Users\\Admin\\data.xlsx");
        var (type, message, _) = SensitiveDataRedactor.RedactException(exception);

        output.WriteLine($"Exception Type: {type}");
        output.WriteLine($"Redacted Message: {message}");

        Assert.Equal("InvalidOperationException", type);
        Assert.DoesNotContain("C:\\", message);
        Assert.Contains("[REDACTED_PATH]", message);
    }

    [Fact]
    public void ToolInvocation_ExecutesWithTelemetry()
    {
        output.WriteLine("=== TOOL INVOCATION TEST ===\n");

        // Act - call a tool method that uses ExecuteToolAction
        // Using Test action since it doesn't require an actual file
        var result = ExcelFileTool.ExcelFile(
            FileAction.Test,
            excelPath: "C:\\fake\\test.xlsx",
            sessionId: null,
            save: false,
            showExcel: false,
            timeoutSeconds: 300);

        output.WriteLine($"Tool result: {result[..Math.Min(200, result.Length)]}...\n");

        // Assert - tool executed (telemetry is tracked internally)
        Assert.NotNull(result);
        Assert.Contains("success", result.ToLowerInvariant());
    }
}
