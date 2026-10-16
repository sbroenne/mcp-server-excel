using Microsoft.Win32;
using System.Globalization;

namespace Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;

/// <summary>
/// Raw Power Query tracing values. A null value means the value is not set.
/// </summary>
/// <param name="EnabledValue">TracingEnabled value (0 or 1), or null when absent.</param>
/// <param name="LevelValue">TracingLevel value, or null when absent.</param>
public sealed record PowerQueryTracingSnapshot(int? EnabledValue, int? LevelValue)
{
    /// <summary>Whether tracing is switched on.</summary>
    public bool Enabled => EnabledValue is int value && value != 0;
}

/// <summary>
/// Reads and writes the per-user Power Query tracing switch.
/// </summary>
public interface IPowerQueryTraceSettings
{
    /// <summary>Folder where Power Query writes trace files.</summary>
    string TraceFolder { get; }

    /// <summary>Reads the current values.</summary>
    PowerQueryTracingSnapshot Read();

    /// <summary>Switches tracing on or off and optionally sets the level.</summary>
    void Write(bool enabled, int? level);
}

/// <summary>
/// Power Query tracing switch stored in
/// <c>HKCU\Software\Microsoft\Office\16.0\PowerQuery</c> (same as Query Options, Diagnostics, Enable tracing).
/// </summary>
public sealed class RegistryPowerQueryTraceSettings : IPowerQueryTraceSettings
{
    internal const string KeyPath = @"Software\Microsoft\Office\16.0\PowerQuery";
    internal const string EnabledValueName = "TracingEnabled";
    internal const string LevelValueName = "TracingLevel";
    internal const int DefaultTracingLevel = 15;

    /// <inheritdoc />
    public string TraceFolder { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Microsoft", "Office", "16.0", "PowerQuery", "Traces");

    /// <inheritdoc />
    public PowerQueryTracingSnapshot Read()
    {
        using var key = Registry.CurrentUser.OpenSubKey(KeyPath, writable: false);
        return new PowerQueryTracingSnapshot(
            ReadInt(key, EnabledValueName),
            ReadInt(key, LevelValueName));
    }

    /// <inheritdoc />
    public void Write(bool enabled, int? level)
    {
        using var key = Registry.CurrentUser.CreateSubKey(KeyPath, writable: true);
        key.SetValue(EnabledValueName, enabled ? 1 : 0, RegistryValueKind.DWord);
        if (level is int value)
        {
            key.SetValue(LevelValueName, value, RegistryValueKind.DWord);
        }
    }

    private static int? ReadInt(RegistryKey? key, string name)
    {
        return key?.GetValue(name) switch
        {
            int number => number,
            string text when int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) => parsed,
            _ => null
        };
    }
}
