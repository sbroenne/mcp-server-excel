using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Calculation;

/// <summary>Calculation modes matching native Excel values.</summary>
public enum CalculationMode
{
    /// <summary>Recalculate when values change.</summary>
    Automatic = -4105,
    /// <summary>Calculate only when requested.</summary>
    Manual = -4135,
    /// <summary>Automatic except what-if data tables.</summary>
    SemiAutomatic = 2
}

/// <summary>Native calculation target.</summary>
public enum CalculationScope
{
    /// <summary>All open workbooks in the session's owned Excel application.</summary>
    Application,
    /// <summary>One worksheet.</summary>
    Sheet,
    /// <summary>An exact worksheet range.</summary>
    Range
}

/// <summary>Native calculation strength.</summary>
public enum CalculationKind
{
    /// <summary>Recalculate dirty formulas.</summary>
    Normal,
    /// <summary>Recalculate every formula in the owned application.</summary>
    Full,
    /// <summary>Rebuild dependencies and recalculate every formula in the owned application.</summary>
    Rebuild
}

/// <summary>Native application settings and session workbook precision.</summary>
public sealed class CalculationSettingsResult : OperationResult
{
    /// <summary>Mode: automatic, manual, or semi-automatic.</summary>
    public string Mode { get; set; } = string.Empty;
    /// <summary>Native calculation mode value.</summary>
    public int ModeValue { get; set; }
    /// <summary>done, calculating, or pending.</summary>
    public string CalculationState { get; set; } = string.Empty;
    /// <summary>Native calculation state value.</summary>
    public int CalculationStateValue { get; set; }
    /// <summary>True unless native calculation state is done.</summary>
    public bool IsPending { get; set; }
    /// <summary>Mode and iteration settings affect the owned application.</summary>
    public string SettingsScope { get; } = "application";
    /// <summary>Whether circular formulas are calculated iteratively.</summary>
    public bool IterationEnabled { get; set; }
    /// <summary>Maximum iteration count.</summary>
    public int MaximumIterations { get; set; }
    /// <summary>Native convergence tolerance.</summary>
    public double MaximumChange { get; set; }
    /// <summary>Whether Excel calculates before saving in manual mode.</summary>
    public bool CalculateBeforeSave { get; set; }
    /// <summary>Session workbook precision-as-displayed flag, not an application setting.</summary>
    public bool PrecisionAsDisplayed { get; set; }
}

/// <summary>
/// Read/change application calculation settings and explicit workbook precision.
/// Preserve the prior mode after temporary bulk-write changes, including failure.
/// Writes attempt restoration; use get-settings when subsequent work depends on it.
/// Full/rebuild calculation applies to all workbooks in the owned application, not other Excel processes.
/// Precision-as-displayed permanently rounds stored values; disabling it does not recover lost digits.
/// </summary>
[ServiceCategory("CalculationMode")]
[McpTool("calculation_mode", Title = "Calculation Settings", Destructive = true, Category = "settings",
    Description = "Change native calculation settings and explicitly recalculate formulas. set-settings replaces set-mode. " +
        "Mode, iteration_enabled, maximum_iterations, maximum_change and calculate_before_save affect the owned Excel application; omitted set-settings inputs stay unchanged. " +
        "Value/formula writes attempt to restore the prior mode; restoration can fail without failing the write. Verify the mode when subsequent work depends on it. " +
        "For costly bulk writes remember the current mode, set-settings mode manual, write, calculate, then restore the prior mode, including after failure. " +
        "Automatic normally recalculates dependent formulas; manual needs explicit calculate. " +
        "Semi-automatic excludes what-if data tables, not worksheet Tables. Successful writes/calculation do not establish completion of asynchronous refreshes or Python calculations. " +
        "calculate requires scope application, sheet or range; sheet/range require sheet_name and range also requires range_address. " +
        "kind normal is the default; full and rebuild require application scope and affect ALL open workbooks in the owned Excel process. No activation/selection or mode changes. " +
        "set-precision affects only the session workbook. precision_as_displayed true requires allow_precision_loss true: stored numeric precision is permanently lost; disabling does not recover digits.")]
[McpReadOnlyActions("get-settings")]
public interface ICalculationModeCommands
{
    /// <summary>Read actual calculation settings/state and workbook precision; no success-shaped fallback.</summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("get-settings")]
    CalculationSettingsResult GetSettings(IExcelBatch batch);

    /// <summary>
    /// Change only supplied application settings and return native readback.
    /// Requires at least one setting; native failure does not promise rollback.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="mode">Optional automatic/manual/semi-automatic mode</param>
    /// <param name="iterationEnabled">Optional circular-formula iteration flag</param>
    /// <param name="maximumIterations">Optional iteration count, 1 through 32767</param>
    /// <param name="maximumChange">Optional finite positive convergence tolerance</param>
    /// <param name="calculateBeforeSave">Optional native calculate-before-save flag</param>
    [ServiceAction("set-settings")]
    CalculationSettingsResult SetSettings(IExcelBatch batch,
        [FromString] CalculationMode? mode = null, bool? iterationEnabled = null,
        int? maximumIterations = null, double? maximumChange = null, bool? calculateBeforeSave = null);

    /// <summary>
    /// Explicitly change session workbook precision-as-displayed and return readback.
    /// Enabling permanently rounds stored values; disabling does not restore digits.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="precisionAsDisplayed">Required true to enable or false to disable</param>
    /// <param name="allowPrecisionLoss">Required explicit true permission when enabling; default false</param>
    [ServiceAction("set-precision")]
    CalculationSettingsResult SetPrecision(IExcelBatch batch,
        [RequiredParameter] bool precisionAsDisplayed, bool allowPrecisionLoss = false);

    /// <summary>
    /// Recalculate the selected native scope; full/rebuild require application scope.
    /// Application calculation affects all workbooks in this owned Excel process.
    /// Does not change mode, selection, or activation, or wait for asynchronous refresh.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="scope">Required application, sheet, or range; former workbook spelling is removed</param>
    /// <param name="sheetName">Required for sheet/range scope</param>
    /// <param name="rangeAddress">Required for range scope</param>
    /// <param name="kind">normal, full, or rebuild; default normal</param>
    [ServiceAction("calculate")]
    OperationResult Calculate(IExcelBatch batch, [RequiredParameter, FromString] CalculationScope scope,
        string? sheetName = null, string? rangeAddress = null,
        [FromString] CalculationKind kind = CalculationKind.Normal);
}

/// <summary>Calculation settings and execution through native Excel.</summary>
public sealed class CalculationModeCommands : ICalculationModeCommands
{
    /// <inheritdoc />
    public CalculationSettingsResult GetSettings(IExcelBatch batch) =>
        batch.Execute((context, ct) =>
        {
            ct.ThrowIfCancellationRequested();
            return ReadSettings(context, batch.WorkbookPath, "get-settings");
        });

    /// <inheritdoc />
    public CalculationSettingsResult SetSettings(IExcelBatch batch, CalculationMode? mode = null,
        bool? iterationEnabled = null, int? maximumIterations = null,
        double? maximumChange = null, bool? calculateBeforeSave = null)
    {
        if (mode.HasValue && !Enum.IsDefined(mode.Value))
            throw new ArgumentOutOfRangeException(nameof(mode), mode, $"Unknown calculation mode: {mode}");
        if (maximumIterations is < 1 or > 32767)
            throw new ArgumentOutOfRangeException(nameof(maximumIterations), "Iteration count must be 1 through 32767.");
        if (maximumChange.HasValue && (!double.IsFinite(maximumChange.Value) || maximumChange.Value <= 0))
            throw new ArgumentOutOfRangeException(nameof(maximumChange), "Convergence tolerance must be finite and positive.");
        if (!mode.HasValue && !iterationEnabled.HasValue && !maximumIterations.HasValue &&
            !maximumChange.HasValue && !calculateBeforeSave.HasValue)
            throw new ArgumentException("Supply at least one calculation setting.");
        return batch.Execute((context, ct) =>
        {
            ct.ThrowIfCancellationRequested();
            if (maximumIterations.HasValue) context.App.MaxIterations = maximumIterations.Value;
            if (maximumChange.HasValue) context.App.MaxChange = maximumChange.Value;
            if (iterationEnabled.HasValue) context.App.Iteration = iterationEnabled.Value;
            if (calculateBeforeSave.HasValue) context.App.CalculateBeforeSave = calculateBeforeSave.Value;
            if (mode.HasValue) context.App.Calculation = (Excel.XlCalculation)mode.Value;
            return ReadSettings(context, batch.WorkbookPath, "set-settings");
        });
    }

    /// <inheritdoc />
    public CalculationSettingsResult SetPrecision(IExcelBatch batch, bool precisionAsDisplayed,
        bool allowPrecisionLoss = false)
    {
        if (precisionAsDisplayed && !allowPrecisionLoss)
            throw new InvalidOperationException("Enabling precision-as-displayed permanently loses stored numeric precision. Explicit allowPrecisionLoss=true is required.");
        return batch.Execute((context, ct) =>
        {
            ct.ThrowIfCancellationRequested();
            context.Book.PrecisionAsDisplayed = precisionAsDisplayed;
            var result = ReadSettings(context, batch.WorkbookPath, "set-precision");
            result.Message = precisionAsDisplayed
                ? "Precision-as-displayed enabled for this workbook; stored numeric precision may be permanently lost."
                : "Precision-as-displayed disabled; previously lost digits are not recovered.";
            return result;
        });
    }

    /// <inheritdoc />
    public OperationResult Calculate(IExcelBatch batch, CalculationScope scope,
        string? sheetName = null, string? rangeAddress = null, CalculationKind kind = CalculationKind.Normal)
    {
        if (!Enum.IsDefined(scope))
            throw new ArgumentOutOfRangeException(nameof(scope), scope, $"Unknown calculation scope: {scope}");
        if (!Enum.IsDefined(kind))
            throw new ArgumentOutOfRangeException(nameof(kind));
        if (scope != CalculationScope.Application && kind != CalculationKind.Normal)
            throw new ArgumentException("Full/rebuild calculation requires application scope.", nameof(kind));
        if (scope != CalculationScope.Application && string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("sheetName is required for Sheet/Range scope calculation.", nameof(sheetName));
        if (scope == CalculationScope.Range && string.IsNullOrWhiteSpace(rangeAddress))
            throw new ArgumentException("rangeAddress is required for Range scope calculation.", nameof(rangeAddress));
        if (scope == CalculationScope.Application && !string.IsNullOrEmpty(sheetName) ||
            scope != CalculationScope.Range && !string.IsNullOrEmpty(rangeAddress))
            throw new ArgumentException("Sheet/range inputs must match the requested calculation scope.");
        return batch.Execute((context, ct) =>
        {
            ct.ThrowIfCancellationRequested();
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                if (scope == CalculationScope.Application)
                {
                    switch (kind)
                    {
                        case CalculationKind.Normal: context.App.Calculate(); break;
                        case CalculationKind.Full: context.App.CalculateFull(); break;
                        case CalculationKind.Rebuild: context.App.CalculateFullRebuild(); break;
                        default: throw new ArgumentOutOfRangeException(nameof(kind));
                    }
                }
                else
                {
                    sheet = ComUtilities.FindSheet(context.Book, sheetName!);
                    if (sheet is null)
                        throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
                    if (scope == CalculationScope.Sheet)
                        sheet.Calculate();
                    else
                    {
                        range = Range.RangeHelpers.ResolveRange(context.Book, sheetName!, rangeAddress!, out string? error);
                        if (range is null)
                            throw new InvalidOperationException(error ?? $"Range '{rangeAddress}' was not found.");
                        range.Calculate();
                    }
                }
                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "calculate",
                    Message = $"{kind} calculation completed for {scope}; asynchronous refresh/Python completion is not established."
                };
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static CalculationSettingsResult ReadSettings(ExcelContext context, string filePath, string action)
    {
        var mode = context.App.Calculation;
        var state = context.App.CalculationState;
        return new CalculationSettingsResult
        {
            Success = true,
            FilePath = filePath,
            Action = action,
            Mode = mode switch
            {
                Excel.XlCalculation.xlCalculationAutomatic => "automatic",
                Excel.XlCalculation.xlCalculationManual => "manual",
                Excel.XlCalculation.xlCalculationSemiautomatic => "semi-automatic",
                _ => throw new InvalidOperationException($"Unknown native calculation mode: {(int)mode}.")
            },
            ModeValue = (int)mode,
            CalculationState = state switch
            {
                Excel.XlCalculationState.xlDone => "done",
                Excel.XlCalculationState.xlCalculating => "calculating",
                Excel.XlCalculationState.xlPending => "pending",
                _ => throw new InvalidOperationException($"Unknown native calculation state: {(int)state}.")
            },
            CalculationStateValue = (int)state,
            IsPending = state != Excel.XlCalculationState.xlDone,
            IterationEnabled = context.App.Iteration,
            MaximumIterations = context.App.MaxIterations,
            MaximumChange = context.App.MaxChange,
            CalculateBeforeSave = context.App.CalculateBeforeSave,
            PrecisionAsDisplayed = context.Book.PrecisionAsDisplayed
        };
    }
}
