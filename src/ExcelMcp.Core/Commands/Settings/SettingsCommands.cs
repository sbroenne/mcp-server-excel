using Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Settings;

/// <summary>
/// Per-user Excel settings commands.
/// </summary>
public sealed class SettingsCommands : ISettingsCommands
{
    internal const string NativeQueryApprovalWarning =
        "This setting applies to all workbooks for the current Windows user. Turning approval off allows native database queries to run without this approval.";

    internal const string TracingTakesEffectNote =
        "This is a per-user setting, not a workbook setting. Restart Excel for the change to take effect. Any Excel process started while tracing is on may write trace files.";

    private readonly IPowerQueryTraceSettings _traceSettings;
    private readonly INativeQueryApprovalSettings _nativeQueryApprovalSettings;
    private readonly Func<bool> _isExcelRunning;

    /// <summary>Creates settings commands that use the current user's registry.</summary>
    public SettingsCommands(IPowerQueryTraceSettings? traceSettings = null) =>
        (_traceSettings, _nativeQueryApprovalSettings, _isExcelRunning) = (
            traceSettings ?? new RegistryPowerQueryTraceSettings(),
            new NativeQueryApprovalSettings(),
            NativeQueryApprovalSettings.IsExcelRunning);

    internal SettingsCommands(
        IPowerQueryTraceSettings? traceSettings,
        INativeQueryApprovalSettings nativeQueryApprovalSettings,
        Func<bool> isExcelRunning)
    {
        _traceSettings = traceSettings ?? new RegistryPowerQueryTraceSettings();
        _nativeQueryApprovalSettings = nativeQueryApprovalSettings;
        _isExcelRunning = isExcelRunning;
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

        var previous = _traceSettings.Read();
        _traceSettings.Write(enabled, level);
        var current = _traceSettings.Read();
        var result = CreateStatus(current);
        result.PreviousEnabled = previous.Enabled;
        result.PreviousLevel = previous.LevelValue ?? RegistryPowerQueryTraceSettings.DefaultTracingLevel;
        return result;
    }

    /// <inheritdoc />
    public NativeQueryApprovalResult GetNativeQueryApproval() =>
        CreateNativeQueryApprovalResult(_nativeQueryApprovalSettings.Read());

    /// <inheritdoc />
    public NativeQueryApprovalResult SetNativeQueryApproval(bool? required, bool confirm)
    {
        if (!confirm)
        {
            throw new InvalidOperationException(
                "Changing the global native-query approval setting requires confirm=true. This setting applies to all workbooks for the current Windows user.");
        }

        if (required is null)
        {
            throw new InvalidOperationException(
                "Pass required=true or required=false. The approval requirement must not be omitted.");
        }

        bool requestedRequired = required.Value;
        NativeQueryApprovalSnapshot previous = _nativeQueryApprovalSettings.Read();
        if (previous.Required == requestedRequired)
        {
            NativeQueryApprovalResult unchanged = CreateNativeQueryApprovalResult(previous);
            unchanged.PreviousRequired = previous.Required;
            return unchanged;
        }

        if (_isExcelRunning())
        {
            throw new InvalidOperationException(
                "Close all Excel windows before changing native-query approval. No Excel process was stopped and the setting was not changed.");
        }

        string backupPath = _nativeQueryApprovalSettings.Write(requestedRequired);
        NativeQueryApprovalSnapshot current;
        try
        {
            current = _nativeQueryApprovalSettings.Read();
        }
        catch (Exception exception)
        {
            throw new InvalidOperationException(
                $"The native-query approval archive was replaced, but the new setting could not be verified. The backup is '{backupPath}'.",
                exception);
        }

        if (current.Required != requestedRequired || !current.ExplicitlySet)
        {
            throw new InvalidOperationException(
                $"The native-query approval archive was replaced, but verification returned required={current.Required} and explicitlySet={current.ExplicitlySet}, not required={requestedRequired}. The backup is '{backupPath}'.");
        }

        NativeQueryApprovalResult result = CreateNativeQueryApprovalResult(current);
        result.PreviousRequired = previous.Required;
        result.BackupPath = backupPath;
        return result;
    }

    private NativeQueryApprovalResult CreateNativeQueryApprovalResult(NativeQueryApprovalSnapshot snapshot) =>
        new()
        {
            Success = true,
            Required = snapshot.Required,
            ExplicitlySet = snapshot.ExplicitlySet,
            SettingsFilePath = _nativeQueryApprovalSettings.SettingsFilePath,
            Warning = NativeQueryApprovalWarning
        };

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
