using Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Settings;

/// <summary>
/// Per-user Excel settings commands.
/// </summary>
public sealed class SettingsCommands : ISettingsCommands
{
    internal const string TracingTakesEffectNote =
        "This is a per-user setting, not a workbook setting. Restart Excel for the change to take effect. Any Excel process started while tracing is on may write trace files.";

    private readonly IPowerQueryTraceSettings _traceSettings;

    /// <summary>Creates settings commands that use the current user's registry.</summary>
    public SettingsCommands(IPowerQueryTraceSettings? traceSettings = null) =>
        _traceSettings = traceSettings ?? new RegistryPowerQueryTraceSettings();

    /// <inheritdoc />
    public PowerQueryTracingStatusResult GetPqTracing()
    {
        var current = _traceSettings.Read();
        return CreateStatus(current);
    }

    /// <inheritdoc />
    public PowerQueryTracingStatusResult SetPqTracing(bool enabled, int? level = null)
    {
        if (level is < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(level), level, "level must be zero or greater.");
        }

        var previous = _traceSettings.Read();
        _traceSettings.Write(enabled, level);
        var current = _traceSettings.Read();
        var result = CreateStatus(current);
        result.PreviousEnabled = previous.Enabled;
        result.PreviousLevel = previous.LevelValue ?? RegistryPowerQueryTraceSettings.DefaultTracingLevel;
        return result;
    }

    private PowerQueryTracingStatusResult CreateStatus(PowerQueryTracingSnapshot snapshot)
    {
        return new PowerQueryTracingStatusResult
        {
            Success = true,
            Enabled = snapshot.Enabled,
            Level = snapshot.LevelValue ?? RegistryPowerQueryTraceSettings.DefaultTracingLevel,
            TraceFolder = _traceSettings.TraceFolder,
            TraceFileCount = PowerQueryTraceFiles.Count(_traceSettings.TraceFolder),
            Note = TracingTakesEffectNote
        };
    }
}
