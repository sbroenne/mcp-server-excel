// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

/// <summary>
/// Tests for telemetry configuration and sensitive data redaction.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Telemetry")]
public class TelemetryTests
{
    #region ExcelMcpTelemetry Tests

    [Fact]
    public void SessionId_IsNotEmpty()
    {
        // Session ID should be generated on startup
        Assert.False(string.IsNullOrEmpty(ExcelMcpTelemetry.SessionId));
    }

    [Fact]
    public void SessionId_IsEightCharacters()
    {
        // Session ID should be first 8 chars of GUID
        Assert.Equal(8, ExcelMcpTelemetry.SessionId.Length);
    }

    [Fact]
    public void UserId_IsNotEmpty()
    {
        // User ID should be generated from machine identity
        Assert.False(string.IsNullOrEmpty(ExcelMcpTelemetry.UserId));
    }

    [Fact]
    public void UserId_IsSixteenCharacters()
    {
        // User ID should be first 16 chars of SHA256 hash
        Assert.Equal(16, ExcelMcpTelemetry.UserId.Length);
    }

    [Fact]
    public void UserId_IsLowercaseHex()
    {
        // User ID should be lowercase hex characters only
        Assert.True(ExcelMcpTelemetry.UserId.All(c => char.IsAsciiHexDigitLower(c)));
    }

    [Fact]
    public void GetConnectionString_ReturnsNullForPlaceholder()
    {
        // The placeholder should not be treated as a valid connection string
        // (In dev builds, it's "__APPINSIGHTS_CONNECTION_STRING__")
        var connectionString = ExcelMcpTelemetry.GetConnectionString();

        // Either null (placeholder) or a real connection string (CI build)
        // We can't assert null directly because CI might inject a real one
        if (connectionString != null)
        {
            Assert.DoesNotContain("__", connectionString);
        }
    }

    #endregion

    #region SensitiveDataRedactor Tests

    [Theory]
    [InlineData(@"C:\Users\john\Documents\file.xlsx", "[REDACTED_PATH]")]
    [InlineData(@"D:\source\project\data.csv", "[REDACTED_PATH]")]
    [InlineData(@"E:\folder\subfolder\test.txt", "[REDACTED_PATH]")]
    public void RedactSensitiveData_RedactsWindowsPaths(string input, string expected)
    {
        var result = SensitiveDataRedactor.RedactSensitiveData(input);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData(@"\\server\share\file.xlsx", "[REDACTED_PATH]")]
    [InlineData(@"\\192.168.1.1\data\report.csv", "[REDACTED_PATH]")]
    [InlineData(@"\\company.local\shared\docs\file.txt", "[REDACTED_PATH]")]
    public void RedactSensitiveData_RedactsUncPaths(string input, string expected)
    {
        var result = SensitiveDataRedactor.RedactSensitiveData(input);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("Password=secret123", "Password=[REDACTED]")]
    [InlineData("pwd=mypassword", "pwd=[REDACTED]")]
    [InlineData("User Id=admin;Password=secret", "User Id=admin;Password=[REDACTED]")]
    public void RedactSensitiveData_RedactsPasswords(string input, string expected)
    {
        var result = SensitiveDataRedactor.RedactSensitiveData(input);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("user@example.com", "[REDACTED_EMAIL]")]
    [InlineData("john.doe@company.org", "[REDACTED_EMAIL]")]
    [InlineData("Contact: admin@test.co.uk for help", "Contact: [REDACTED_EMAIL] for help")]
    public void RedactSensitiveData_RedactsEmails(string input, string expected)
    {
        var result = SensitiveDataRedactor.RedactSensitiveData(input);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("https://user:pass@server.com/api", "https://[REDACTED]@server.com/api")]
    [InlineData("http://admin:secret123@localhost:8080", "http://[REDACTED]@localhost:8080")]
    public void RedactSensitiveData_RedactsUrlCredentials(string input, string expected)
    {
        var result = SensitiveDataRedactor.RedactSensitiveData(input);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("Operation completed successfully", "Operation completed successfully")]
    [InlineData("Error code 500", "Error code 500")]
    [InlineData("Range A1:B10", "Range A1:B10")]
    public void RedactSensitiveData_PreservesNonSensitiveData(string input, string expected)
    {
        var result = SensitiveDataRedactor.RedactSensitiveData(input);
        Assert.Equal(expected, result);
    }

    [Fact]
    public void RedactSensitiveData_HandlesEmptyInput()
    {
        var result = SensitiveDataRedactor.RedactSensitiveData(string.Empty);
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void RedactSensitiveData_RedactsMultipleSensitiveItems()
    {
        var input = @"Error accessing C:\Users\john\file.xlsx: user@example.com failed with Password=secret";
        var result = SensitiveDataRedactor.RedactSensitiveData(input);

        Assert.DoesNotContain(@"C:\Users", result);
        Assert.DoesNotContain("user@example.com", result);
        Assert.DoesNotContain("Password=secret", result);
        Assert.Contains("[REDACTED_PATH]", result);
        Assert.Contains("[REDACTED_EMAIL]", result);
        Assert.Contains("[REDACTED]", result);
    }

    [Fact]
    public void RedactException_RedactsExceptionMessage()
    {
        var exception = new InvalidOperationException(@"Failed to open C:\Users\admin\secret.xlsx");

        var (type, message, _) = SensitiveDataRedactor.RedactException(exception);

        Assert.Equal("InvalidOperationException", type);
        Assert.Contains("[REDACTED_PATH]", message);
        Assert.DoesNotContain(@"C:\Users", message);
    }

    [Fact]
    public void RedactException_PreservesExceptionType()
    {
        var exception = new ArgumentException("Test error");

        var (type, message, _) = SensitiveDataRedactor.RedactException(exception);

        Assert.Equal("ArgumentException", type);
        Assert.Equal("Test error", message);
    }

    [Fact]
    public void RedactException_RedactsStackTrace()
    {
        InvalidOperationException caughtException;
        try
        {
            // Create exception with stack trace containing path
            throw new InvalidOperationException(@"Error at C:\Users\test\file.cs line 42");
        }
        catch (InvalidOperationException ex)
        {
            caughtException = ex;
        }

        var (_, message, stackTrace) = SensitiveDataRedactor.RedactException(caughtException);

        Assert.Contains("[REDACTED_PATH]", message);
        // Stack trace will contain the actual test file path which should be redacted
        if (stackTrace != null)
        {
            // The stack trace contains this test file's path
            Assert.DoesNotContain(@"C:\Users", stackTrace);
        }
    }

    #endregion
}




