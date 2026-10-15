using System.Diagnostics;
using System.Security;
using System.Security.Cryptography;
using System.Text;

namespace Sbroenne.ExcelMcp.Telemetry;

internal sealed record TelemetryUserIdentity(string UserId, string? LegacyUserId)
{
    internal string RoleInstance => $"instance-{UserId[..8]}";
}

internal static class TelemetryUserIdentityStore
{
    internal const string IdentityFileName = "telemetry-user-id";
    private const string ApplicationDirectoryName = "ExcelMcp";

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Optional telemetry storage failures are traced without private path details and fall back to a process-only random identity; product work must not depend on telemetry identity persistence.")]
    internal static TelemetryUserIdentity GetOrCreate()
    {
        try
        {
            var localAppDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return string.IsNullOrWhiteSpace(localAppDataPath)
                ? CreateProcessOnly()
                : GetOrCreate(localAppDataPath);
        }
        catch (Exception exception) when (IsStorageFailure(exception))
        {
            TraceStorageFailure(exception);
            return CreateProcessOnly();
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Optional telemetry storage failures are traced without private path details and fall back to a process-only random identity; product work must not depend on telemetry identity persistence.")]
    internal static TelemetryUserIdentity GetOrCreate(string localAppDataPath)
    {
        string? temporaryPath = null;
        try
        {
            var directoryPath = Path.Combine(localAppDataPath, ApplicationDirectoryName);
            Directory.CreateDirectory(directoryPath);
            var identityPath = Path.Combine(directoryPath, IdentityFileName);

            if (File.Exists(identityPath))
            {
                return TryRead(identityPath) ?? CreateProcessOnly();
            }

            var userId = Guid.NewGuid().ToString("N");
            temporaryPath = Path.Combine(directoryPath, $"{IdentityFileName}.{Guid.NewGuid():N}.tmp");
            using (var stream = new FileStream(
                       temporaryPath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.None))
            {
                var bytes = Encoding.UTF8.GetBytes(userId);
                stream.Write(bytes);
                stream.Flush(flushToDisk: true);
            }

            try
            {
                File.Move(temporaryPath, identityPath);
                temporaryPath = null;
                return new TelemetryUserIdentity(userId, GenerateLegacyUserId());
            }
            catch (IOException) when (File.Exists(identityPath))
            {
                return TryRead(identityPath) ?? CreateProcessOnly();
            }
        }
        catch (Exception exception) when (IsStorageFailure(exception))
        {
            TraceStorageFailure(exception);
            return CreateProcessOnly();
        }
        finally
        {
            if (temporaryPath is not null)
            {
                TryDeleteTemporaryFile(temporaryPath);
            }
        }
    }

    internal static TelemetryUserIdentity CreateProcessOnly() =>
        new(Guid.NewGuid().ToString("N"), null);

    internal static UsageEvent? CreateMigrationEvent(TelemetryUserIdentity identity)
    {
        if (identity.LegacyUserId is null)
        {
            return null;
        }

        var telemetry = new UsageEvent("UserIdMigration");
        telemetry.Properties["LegacyUserId"] = identity.LegacyUserId;
        return telemetry;
    }

    private static TelemetryUserIdentity? TryRead(string identityPath)
    {
        var storedValue = File.ReadAllText(identityPath);
        if (Guid.TryParseExact(storedValue, "N", out var storedId))
        {
            return new TelemetryUserIdentity(storedId.ToString("N"), null);
        }

        Trace.TraceWarning("Telemetry user ID storage is invalid; using a process-only ID.");
        return null;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Only expected platform/security failures while constructing an optional migration identifier are handled; the failure type is traced and no migration event is emitted.")]
    private static string? GenerateLegacyUserId()
    {
        try
        {
            var machineIdentity =
                $"{Environment.MachineName}|{Environment.UserName}|{Environment.OSVersion.Platform}";
            var hash = SHA256.HashData(Encoding.UTF8.GetBytes(machineIdentity));
            return Convert.ToHexString(hash)[..16].ToLowerInvariant();
        }
        catch (Exception exception) when (
            exception is InvalidOperationException or SecurityException or PlatformNotSupportedException)
        {
            Trace.TraceWarning("Telemetry migration ID could not be generated ({0}).", exception.GetType().Name);
            return null;
        }
    }

    private static bool IsStorageFailure(Exception exception) =>
        exception is IOException or UnauthorizedAccessException or ArgumentException
            or NotSupportedException or SecurityException;

    private static void TraceStorageFailure(Exception exception) =>
        Trace.TraceWarning(
            "Telemetry user ID storage is unavailable; using a process-only ID ({0}).",
            exception.GetType().Name);

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Telemetry identity scratch-file cleanup handles only expected storage failures and traces the exception type without exposing private paths; it must not fail product work.")]
    internal static void TryDeleteTemporaryFile(string path, Action<string>? delete = null)
    {
        try
        {
            (delete ?? File.Delete)(path);
        }
        catch (Exception exception) when (IsStorageFailure(exception))
        {
            Trace.TraceWarning("A temporary telemetry file could not be removed ({0}).", exception.GetType().Name);
        }
    }
}
