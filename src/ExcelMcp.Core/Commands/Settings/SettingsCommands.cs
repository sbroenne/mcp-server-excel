using Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Settings;

/// <summary>
/// Per-user Excel settings commands.
/// </summary>
public sealed partial class SettingsCommands : ISettingsCommands
{
    internal const string TracingTakesEffectNote =
        "This is a per-user setting, not a workbook setting. Any Excel process started while tracing is on may write trace files; processes already running keep their current tracing state.";

    private readonly IPowerQueryTraceSettings _traceSettings;
    private readonly PowerQueryTracingLeases _tracingLeases;

    /// <summary>Creates settings commands that use the current user's registry.</summary>
    public SettingsCommands(
        IPowerQueryTraceSettings? traceSettings = null,
        PowerQueryTracingLeases? tracingLeases = null)
    {
        _traceSettings = traceSettings ?? new RegistryPowerQueryTraceSettings();
        _tracingLeases = tracingLeases ?? new PowerQueryTracingLeases(_traceSettings);
    }

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

        var (previous, current) = _tracingLeases.SetExplicitly(enabled, level);
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
