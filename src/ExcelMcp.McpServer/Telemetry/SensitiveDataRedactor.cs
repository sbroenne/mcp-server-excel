// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.ExcelMcp.ComInterop.Diagnostics;

namespace Sbroenne.ExcelMcp.McpServer.Telemetry;

/// <summary>
/// Utility class that redacts sensitive data from telemetry before it's sent.
/// Removes file paths, connection strings, credentials, and other PII.
/// </summary>
public static class SensitiveDataRedactor
{
    /// <summary>
    /// Redacts all sensitive data from the given string.
    /// </summary>
    public static string RedactSensitiveData(string value) =>
        SensitiveDataSanitizer.Redact(value) ?? string.Empty;

    /// <summary>
    /// Redacts sensitive data from an exception for safe logging.
    /// Returns exception type, redacted message, and redacted stack trace.
    /// </summary>
    public static (string Type, string Message, string? StackTrace) RedactException(Exception ex) =>
        SensitiveDataSanitizer.RedactException(ex);
}


