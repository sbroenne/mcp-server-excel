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

    /// <summary>Puts back previously read values, removing values that were absent.</summary>
    void Restore(PowerQueryTracingSnapshot snapshot);
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

    /// <inheritdoc />
    public void Restore(PowerQueryTracingSnapshot snapshot)
    {
        using var key = Registry.CurrentUser.CreateSubKey(KeyPath, writable: true);
        RestoreValue(key, EnabledValueName, snapshot.EnabledValue);
        RestoreValue(key, LevelValueName, snapshot.LevelValue);
    }

    private static void RestoreValue(RegistryKey key, string name, int? value)
    {
        if (value is int number)
        {
            key.SetValue(name, number, RegistryValueKind.DWord);
        }
        else
        {
            key.DeleteValue(name, throwOnMissingValue: false);
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

/// <summary>
/// Switches Power Query tracing on for sessions that ask for it and puts the user's
/// original setting back when the last of those sessions closes.
/// </summary>
public sealed class PowerQueryTracingLeases
{
    private readonly IPowerQueryTraceSettings _settings;
    private readonly HashSet<string> _owners = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _lock = new();
    private PowerQueryTracingSnapshot? _original;
    private FileStream? _settingsLock;

    /// <summary>Creates the lease tracker.</summary>
    public PowerQueryTracingLeases(IPowerQueryTraceSettings settings)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
    }

    /// <summary>Switches tracing on for <paramref name="owner"/>.</summary>
    public void Acquire(string owner)
    {
        lock (_lock)
        {
            FileStream? pendingLock = null;
            try
            {
                if (_owners.Count == 0)
                {
                    pendingLock = OpenSettingsLock();
                }

                var current = _settings.Read();
                if (_owners.Count == 0)
                {
                    _original = current;
                }

                if (!current.Enabled)
                {
                    _settings.Write(enabled: true, level: null);
                }

                _owners.Add(owner);
                _settingsLock ??= pendingLock;
                pendingLock = null;
            }
            finally
            {
                pendingLock?.Dispose();
            }
        }
    }

    /// <summary>Moves a lease to a new owner name, for example from a pending key to the session ID.</summary>
    public void Rename(string owner, string newOwner)
    {
        lock (_lock)
        {
            if (_owners.Remove(owner))
            {
                _owners.Add(newOwner);
            }
        }
    }

    /// <summary>
    /// Ends the lease for <paramref name="owner"/>. When no leases remain, the original setting is
    /// put back unless someone switched tracing off in the meantime.
    /// </summary>
    /// <returns>True when the original setting was put back.</returns>
    public bool Release(string owner)
    {
        lock (_lock)
        {
            if (!_owners.Remove(owner) || _owners.Count > 0)
            {
                return false;
            }

            try
            {
                return RestoreOriginal();
            }
            finally
            {
                CloseSettingsLock();
            }
        }
    }

    /// <summary>Ends every lease and puts back the original setting.</summary>
    public void ReleaseAll()
    {
        lock (_lock)
        {
            if (_owners.Count == 0)
            {
                return;
            }

            _owners.Clear();
            try
            {
                RestoreOriginal();
            }
            finally
            {
                CloseSettingsLock();
            }
        }
    }

    internal (PowerQueryTracingSnapshot Previous, PowerQueryTracingSnapshot Current) SetExplicitly(bool enabled, int? level)
    {
        lock (_lock)
        {
            using var settingsLock = _owners.Count == 0 ? OpenSettingsLock() : null;
            var previous = _settings.Read();
            _settings.Write(enabled, level);
            var current = _settings.Read();
            if (_owners.Count > 0)
            {
                _original = current;
            }

            return (previous, current);
        }
    }

    private FileStream OpenSettingsLock()
    {
        Directory.CreateDirectory(_settings.TraceFolder);
        try
        {
            return new FileStream(
                Path.Combine(_settings.TraceFolder, ".excelmcp-tracing.lock"),
                FileMode.OpenOrCreate,
                FileAccess.ReadWrite,
                FileShare.None);
        }
        catch (IOException ex) when ((ex.HResult & 0xffff) is 32 or 33)
        {
            throw new IOException(
                "Another ExcelMcp service is controlling Power Query tracing. Close its tracing-enabled sessions before changing tracing or opening another tracing-enabled session.",
                ex);
        }
    }

    private void CloseSettingsLock()
    {
        _settingsLock?.Dispose();
        _settingsLock = null;
    }

    private bool RestoreOriginal()
    {
        var original = _original;
        _original = null;
        if (original is null || original.Enabled)
        {
            return false;
        }

        // Someone switched tracing off while the lease was held; keep their choice.
        if (!_settings.Read().Enabled)
        {
            return false;
        }

        _settings.Restore(original);
        return true;
    }
}
