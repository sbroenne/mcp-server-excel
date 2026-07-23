// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.McpServer.Telemetry;

/// <summary>
/// Utility class that redacts sensitive data from telemetry before it's sent.
/// Removes file paths, connection strings, credentials, and other PII.
/// </summary>
public static partial class SensitiveDataRedactor
{
    // Patterns for sensitive data detection
    private static readonly Regex FilePathPattern = CreateFilePathRegex();
    private static readonly Regex UncPathPattern = CreateUncPathRegex();
    private static readonly Regex ConnectionStringSecretPattern = CreateConnectionStringSecretRegex();
    private static readonly Regex CredentialPattern = CreateCredentialRegex();
    private static readonly Regex EmailPattern = CreateEmailRegex();

    // Redaction markers
    private const string RedactedPath = "[REDACTED_PATH]";
    private const string RedactedSecret = "[REDACTED]";
    private const string RedactedEmail = "[REDACTED_EMAIL]";

    /// <summary>
    /// Redacts all sensitive data from the given string.
    /// </summary>
    public static string RedactSensitiveData(string value)
    {
        if (string.IsNullOrEmpty(value))
            return value;

        var result = value;

        // Redact file paths (Windows drive letters)
        result = FilePathPattern.Replace(result, RedactedPath);

        // Redact UNC paths
        result = UncPathPattern.Replace(result, RedactedPath);

        // Redact connection string secrets (Password=, pwd=, secret=, key=, token=)
        result = ConnectionStringSecretPattern.Replace(result, match =>
            $"{match.Groups[1].Value}={RedactedSecret}");

        // Redact credentials in URLs (user:pass@host)
        result = CredentialPattern.Replace(result, match =>
            $"{match.Groups[1].Value}{RedactedSecret}@{match.Groups[2].Value}");

        // Redact email addresses
        result = EmailPattern.Replace(result, RedactedEmail);

        return result;
    }

    /// <summary>
    /// Redacts sensitive data from an exception for safe logging.
    /// Returns exception type, redacted message, and redacted stack trace.
    /// </summary>
    public static (string Type, string Message, string? StackTrace) RedactException(Exception ex)
    {
        var type = ex.GetType().Name;
        var message = RedactSensitiveData(ex.Message);
        var stackTrace = ex.StackTrace != null ? RedactSensitiveData(ex.StackTrace) : null;

        return (type, message, stackTrace);
    }

    // Source-generated regex for better performance

    [GeneratedRegex(@"[A-Za-z]:\\[^\s""'<>|*?\r\n]+", RegexOptions.Compiled)]
    private static partial Regex CreateFilePathRegex();

    [GeneratedRegex(@"\\\\[^\s""'<>|*?\r\n]+", RegexOptions.Compiled)]
    private static partial Regex CreateUncPathRegex();

    [GeneratedRegex(@"(Password|pwd|secret|key|token|apikey|api_key|access_token|connectionstring)\s*=\s*[^;""'\s]+", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
    private static partial Regex CreateConnectionStringSecretRegex();

    [GeneratedRegex(@"(https?://)[^:]+:[^@]+@([^\s/]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
    private static partial Regex CreateCredentialRegex();

    [GeneratedRegex(@"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", RegexOptions.Compiled)]
    private static partial Regex CreateEmailRegex();
}


