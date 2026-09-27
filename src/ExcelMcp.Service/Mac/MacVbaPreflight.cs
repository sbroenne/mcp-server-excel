using System.Diagnostics;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal enum MacMacroExecutionAvailability
{
    Available,
    Disabled,
    UserApprovalRequired,
    Unknown
}

internal enum MacVbaProjectModelAccess
{
    Enabled,
    Disabled,
    Unknown
}

internal sealed record MacVbaPreflightResult(
    MacMacroExecutionAvailability MacroExecution,
    MacVbaProjectModelAccess ProjectModelAccess);

internal static class MacVbaPreflight
{
    private const string OfficePreferencesDomain = "com.microsoft.office";
    private static readonly string[] PreferenceKeys =
    [
        "VisualBasicEntirelyDisabled",
        "VisualBasicMacroExecutionState",
        "VBAObjectModelIsTrusted"
    ];

    public static MacVbaPreflightResult Check()
    {
        if (!OperatingSystem.IsMacOS())
        {
            return new MacVbaPreflightResult(
                MacMacroExecutionAvailability.Unknown,
                MacVbaProjectModelAccess.Unknown);
        }

        var preferences = PreferenceKeys.ToDictionary(
            key => key,
            ReadPreference,
            StringComparer.Ordinal);
        return Evaluate(preferences);
    }

    internal static MacVbaPreflightResult Evaluate(
        IReadOnlyDictionary<string, string?> preferences)
    {
        var entirelyDisabled = IsTrue(GetValue(preferences, "VisualBasicEntirelyDisabled"));
        var macroSetting = GetValue(preferences, "VisualBasicMacroExecutionState");
        var macroExecution = entirelyDisabled
            ? MacMacroExecutionAvailability.Disabled
            : macroSetting switch
            {
                "EnabledWithoutWarnings" => MacMacroExecutionAvailability.Available,
                "DisabledWithoutWarnings" => MacMacroExecutionAvailability.Disabled,
                "DisabledWithWarnings" or null or "" =>
                    MacMacroExecutionAvailability.UserApprovalRequired,
                _ => MacMacroExecutionAvailability.Unknown
            };

        var projectTrust = GetValue(preferences, "VBAObjectModelIsTrusted");
        var projectModelAccess = projectTrust is null or ""
            ? MacVbaProjectModelAccess.Disabled
            : IsTrue(projectTrust)
                ? MacVbaProjectModelAccess.Enabled
                : IsFalse(projectTrust)
                    ? MacVbaProjectModelAccess.Disabled
                    : MacVbaProjectModelAccess.Unknown;

        return new MacVbaPreflightResult(macroExecution, projectModelAccess);
    }

    internal static string DescribeMacroExecution(MacMacroExecutionAvailability availability) =>
        availability switch
        {
            MacMacroExecutionAvailability.Available =>
                "Excel is configured for unattended macro execution",
            MacMacroExecutionAvailability.Disabled =>
                "macro execution is disabled by the effective Office preference",
            MacMacroExecutionAvailability.UserApprovalRequired =>
                "the effective Office preference requires per-workbook macro approval",
            _ => "the effective macro execution preference could not be determined"
        };

    internal static string DescribeProjectModel(MacVbaProjectModelAccess access) =>
        access switch
        {
            MacVbaProjectModelAccess.Enabled =>
                "Trust access to the VBA project object model is enabled",
            MacVbaProjectModelAccess.Disabled =>
                "Trust access to the VBA project object model is disabled",
            _ => "Trust access to the VBA project object model could not be determined"
        };

    private static string? GetValue(
        IReadOnlyDictionary<string, string?> preferences,
        string key) =>
        preferences.TryGetValue(key, out var value) ? value?.Trim() : null;

    private static bool IsTrue(string? value) =>
        value is not null
        && (value.Equals("1", StringComparison.Ordinal)
            || value.Equals("true", StringComparison.OrdinalIgnoreCase)
            || value.Equals("yes", StringComparison.OrdinalIgnoreCase));

    private static bool IsFalse(string? value) =>
        value is not null
        && (value.Equals("0", StringComparison.Ordinal)
            || value.Equals("false", StringComparison.OrdinalIgnoreCase)
            || value.Equals("no", StringComparison.OrdinalIgnoreCase));

    private static string? ReadPreference(string key)
    {
        var startInfo = new ProcessStartInfo("/usr/bin/defaults")
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("read");
        startInfo.ArgumentList.Add(OfficePreferencesDomain);
        startInfo.ArgumentList.Add(key);

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Could not start the macOS defaults reader.");
        var output = process.StandardOutput.ReadToEnd();
        if (!process.WaitForExit(1000))
        {
            process.Kill();
            process.WaitForExit();
            throw new TimeoutException("The macOS defaults reader did not complete within one second.");
        }
        return process.ExitCode == 0 ? output.Trim() : null;
    }
}
