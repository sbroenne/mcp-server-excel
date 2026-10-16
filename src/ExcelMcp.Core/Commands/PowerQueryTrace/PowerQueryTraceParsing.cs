using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;

/// <summary>
/// One line of a Power Query trace file.
/// </summary>
internal sealed record PowerQueryTraceEntry(
    string Level,
    string Action,
    string ActivityId,
    DateTime StartUtc,
    TimeSpan Duration,
    int ProcessId,
    int? HostProcessId,
    string? Process,
    string? ResourceKind,
    string? ResourcePath,
    string? CommandText,
    string? Exception,
    bool Cancelled)
{
    public DateTime EndUtc => StartUtc + Duration;
}

/// <summary>
/// Reads Power Query trace lines of the form <c>DataMashup.Trace Level: id : {json}</c>.
/// </summary>
internal static class PowerQueryTraceLineParser
{
    public static bool TryParse(string? line, out PowerQueryTraceEntry? entry)
    {
        entry = null;
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        int jsonStart = line.IndexOf('{');
        if (jsonStart < 0)
        {
            return false;
        }

        string level = ReadLevel(line, jsonStart);

        try
        {
            using var document = JsonDocument.Parse(line.AsMemory(jsonStart));
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            string? action = GetString(root, "Action");
            string? start = GetString(root, "Start");
            if (string.IsNullOrEmpty(action) ||
                start is null ||
                !DateTime.TryParse(start, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var startUtc))
            {
                return false;
            }

            var duration = TimeSpan.Zero;
            string? durationText = GetString(root, "Duration");
            if (durationText != null &&
                (!TimeSpan.TryParse(durationText, CultureInfo.InvariantCulture, out duration) || duration < TimeSpan.Zero))
            {
                return false;
            }

            int? pid = GetInt(root, "Pid");
            if (pid is null)
            {
                return false;
            }

            entry = new PowerQueryTraceEntry(
                level,
                action,
                GetString(root, "ActivityId") ?? string.Empty,
                DateTime.SpecifyKind(startUtc, DateTimeKind.Utc),
                duration,
                pid.Value,
                GetInt(root, "HostProcessId"),
                GetString(root, "Process"),
                GetString(root, "ResourceKind"),
                GetString(root, "ResourcePath"),
                GetString(root, "CommandText"),
                GetString(root, "Exception"),
                GetBool(root, "cancelled") || GetBool(root, "evaluationCancelled"));
            return true;
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static string ReadLevel(string line, int jsonStart)
    {
        // "DataMashup.Trace Information: 24579 : {...}"
        int space = line.IndexOf(' ');
        if (space < 0 || space >= jsonStart)
        {
            return string.Empty;
        }

        int colon = line.IndexOf(':', space);
        return colon > space && colon < jsonStart ? line[(space + 1)..colon] : string.Empty;
    }

    private static string? GetString(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var value))
        {
            return null;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number => value.GetRawText(),
            _ => null
        };
    }

    private static int? GetInt(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var value))
        {
            return null;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out int number))
        {
            return number;
        }

        if (value.ValueKind == JsonValueKind.String &&
            int.TryParse(value.GetString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out number))
        {
            return number;
        }

        return null;
    }

    private static bool GetBool(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var value))
        {
            return false;
        }

        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.String => bool.TryParse(value.GetString(), out bool parsed) && parsed,
            _ => false
        };
    }
}

/// <summary>
/// Reads trace file names such as <c>EXCEL.16772.2026-10-15T07-27-19-661979.log</c>.
/// </summary>
internal static partial class PowerQueryTraceFileName
{
    public const string ExcelKind = "excel";
    public const string ContainerKind = "container";

    [GeneratedRegex(
        @"^(?<process>EXCEL|Microsoft\.Mashup\.Container\.Loader)\.(?<pid>\d+)\.(?<date>\d{4}-\d{2}-\d{2})T(?<hour>\d{2})-(?<minute>\d{2})-(?<second>\d{2})-(?<fraction>\d+)\.log$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex FileNamePattern();

    public static bool TryParse(string fileName, out string kind, out int processId, out DateTime? startedUtc)
    {
        kind = string.Empty;
        processId = 0;
        startedUtc = null;

        var match = FileNamePattern().Match(fileName);
        if (!match.Success ||
            !int.TryParse(match.Groups["pid"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out processId))
        {
            return false;
        }

        kind = match.Groups["process"].Value.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)
            ? ExcelKind
            : ContainerKind;

        string fraction = match.Groups["fraction"].Value;
        fraction = fraction.Length > 7 ? fraction[..7] : fraction;
        string text = $"{match.Groups["date"].Value}T{match.Groups["hour"].Value}:{match.Groups["minute"].Value}:{match.Groups["second"].Value}.{fraction}";
        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var started))
        {
            startedUtc = DateTime.SpecifyKind(started, DateTimeKind.Utc);
        }

        return true;
    }
}

/// <summary>
/// Reads time filters: an ISO 8601 date-time, or an age such as 90s, 15m, 2h, or 7d.
/// </summary>
internal static partial class PowerQueryTraceTime
{
    [GeneratedRegex(@"^(?<amount>\d+)(?<unit>[smhd])$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex AgePattern();

    public static DateTime? Parse(string? value, DateTime nowUtc, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        string text = value.Trim();
        var age = AgePattern().Match(text);
        if (age.Success && long.TryParse(age.Groups["amount"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out long amount))
        {
            try
            {
                var span = char.ToLowerInvariant(age.Groups["unit"].Value[0]) switch
                {
                    's' => TimeSpan.FromSeconds(amount),
                    'm' => TimeSpan.FromMinutes(amount),
                    'h' => TimeSpan.FromHours(amount),
                    _ => TimeSpan.FromDays(amount)
                };
                return nowUtc - span;
            }
            catch (Exception ex) when (ex is OverflowException or ArgumentOutOfRangeException)
            {
                throw new ArgumentException($"{parameterName} '{value}' is too large.", parameterName, ex);
            }
        }

        if (DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsed))
        {
            return parsed.UtcDateTime;
        }

        throw new ArgumentException(
            $"{parameterName} '{value}' is not a valid time. Use an ISO 8601 date-time (for example 2026-01-31T14:05:00Z or a traceCursor value) or an age such as 90s, 15m, 2h, or 7d.",
            parameterName);
    }
}
