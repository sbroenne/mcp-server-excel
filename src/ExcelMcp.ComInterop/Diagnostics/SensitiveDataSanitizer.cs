// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.ComInterop.Diagnostics;

/// <summary>
/// Redacts portable diagnostic text so paths, credentials, connection secrets, and email
/// addresses do not cross service, journal, or telemetry boundaries.
/// </summary>
public static partial class SensitiveDataSanitizer
{
    private static readonly Regex FilePathPattern = CreateFilePathRegex();
    private static readonly Regex UncPathPattern = CreateUncPathRegex();
    private static readonly Regex ConnectionStringSecretPattern = CreateConnectionStringSecretRegex();
    private static readonly Regex CredentialPattern = CreateCredentialRegex();
    private static readonly Regex EmailPattern = CreateEmailRegex();

    /// <summary>
    /// Redacts sensitive fragments from diagnostic text while preserving safe context.
    /// </summary>
    /// <param name="value">Diagnostic text, or <see langword="null"/>.</param>
    /// <returns>The redacted text, preserving <see langword="null"/>.</returns>
    public static string? Redact(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return value;
        }

        var result = FilePathPattern.Replace(value, "[REDACTED_PATH]");
        result = UncPathPattern.Replace(result, "[REDACTED_PATH]");
        result = ConnectionStringSecretPattern.Replace(
            result,
            static match => $"{match.Groups[1].Value}=[REDACTED]");
        result = CredentialPattern.Replace(
            result,
            static match => $"{match.Groups[1].Value}[REDACTED]@{match.Groups[2].Value}");
        return EmailPattern.Replace(result, "[REDACTED_EMAIL]");
    }

    /// <summary>
    /// Extracts an exception's safe type, redacted message, and redacted stack trace.
    /// </summary>
    /// <param name="exception">Exception to describe.</param>
    /// <returns>A portable redacted exception description.</returns>
    public static (string Type, string Message, string? StackTrace) RedactException(Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);
        return (
            exception.GetType().Name,
            Redact(exception.Message) ?? string.Empty,
            Redact(exception.StackTrace));
    }

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
