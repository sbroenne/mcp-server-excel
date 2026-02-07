using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Calculation;

/// <summary>
/// Calculation mode enumeration matching Excel's XlCalculation values.
/// </summary>
public enum CalculationMode
{
    /// <summary>xlCalculationAutomatic - Recalculates when any value changes (default)</summary>
    Automatic = -4105,

    /// <summary>xlCalculationManual - Only recalculates when explicitly requested</summary>
    Manual = -4135,

    /// <summary>xlCalculationSemiautomatic - Auto except data tables (recalc-intensive)</summary>
    SemiAutomatic = 2,
}

/// <summary>
/// Calculation scope for targeted recalculation.
/// </summary>
public enum CalculationScope
{
    /// <summary>Workbook scope - Application.Calculate() recalculates all open workbooks</summary>
    Workbook,

    /// <summary>Sheet scope - Worksheet.Calculate() recalculates single sheet only</summary>
    Sheet,

    /// <summary>Range scope - Range.Calculate() recalculates single range only</summary>
    Range,
}

/// <summary>
/// Result from get-mode action containing current calculation state.
/// </summary>
public class CalculationModeResult : OperationResult
{
    /// <summary>Current calculation mode as string: automatic, manual, semi-automatic</summary>
    public string Mode { get; set; } = string.Empty;

    /// <summary>Raw Excel enumeration value (-4105, -4135, 2)</summary>
    public int ModeValue { get; set; }

    /// <summary>Calculation state (pending, done, etc.)</summary>
    public string CalculationState { get; set; } = string.Empty;

    /// <summary>Raw Excel calculation state value</summary>
    public int CalculationStateValue { get; set; }

    /// <summary>Whether recalculation is pending</summary>
    public bool IsPending { get; set; }

    /// <summary>Sheet name (for sheet/range scope operations)</summary>
    public string? SheetName { get; set; }

    /// <summary>Range address (for range scope operations)</summary>
    public string? RangeAddress { get; set; }

    /// <summary>Calculation scope that was executed</summary>
    public string Scope { get; set; } = string.Empty;
}

/// <summary>
/// Control Excel recalculation (automatic vs manual). Set manual mode before bulk writes
/// for faster performance, then recalculate once at the end.
/// </summary>
[ServiceCategory("calculation", "Calculation")]
[McpTool("excel_calculation_mode", Title = "Excel Calculation Mode Control", Destructive = false, Category = "settings",
    Description = "Set or get Excel calculation mode and explicitly recalculate formulas. MODES: automatic (recalculates on every change, default), manual (only when explicitly requested), semi-automatic (auto except data tables). WORKFLOW for batch operations: 1. set-mode(manual) 2. Perform data operations 3. calculate(workbook) 4. set-mode(automatic). SCOPES for calculate: workbook (all formulas), sheet (requires sheetName), range (requires sheetName + rangeAddress).")]
public interface ICalculationModeCommands
{
    /// <summary>
    /// Gets the current calculation mode and state.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>Current calculation mode (automatic/manual/semi-automatic)</returns>
    [ServiceAction("get-mode")]
    CalculationModeResult GetMode(IExcelBatch batch);

    /// <summary>
    /// Sets the calculation mode (automatic, manual, semi-automatic).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="mode">Target calculation mode</param>
    /// <returns>Operation result with previous and new mode</returns>
    [ServiceAction("set-mode")]
    OperationResult SetMode(IExcelBatch batch, [FromString("mode")] CalculationMode mode);

    /// <summary>
    /// Triggers calculation for the specified scope.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="scope">Scope: Workbook, Sheet, or Range</param>
    /// <param name="sheetName">Sheet name (required for Sheet/Range scope)</param>
    /// <param name="rangeAddress">Range address (required for Range scope)</param>
    /// <returns>Operation result confirming calculation completed</returns>
    [ServiceAction("calculate")]
    OperationResult Calculate(IExcelBatch batch, [FromString("scope")] CalculationScope scope, string? sheetName = null, string? rangeAddress = null);
}

/// <summary>
/// Implementation of calculation mode control commands.
/// </summary>
public class CalculationModeCommands : ICalculationModeCommands
{
    /// <summary>
    /// Gets the current calculation mode and state.
    /// </summary>
    public CalculationModeResult GetMode(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            int modeValue = Convert.ToInt32(ctx.App.Calculation);
            string mode = modeValue switch
            {
                -4105 => "automatic",    // xlCalculationAutomatic
                -4135 => "manual",       // xlCalculationManual
                2 => "semi-automatic",   // xlCalculationSemiautomatic
                _ => "unknown"
            };

            // Get calculation state (if available)
            string calcState = "unknown";
            try
            {
                var calcStateValue = Convert.ToInt32(ctx.App.CalculationState);
                calcState = calcStateValue switch
                {
                    1 => "pending",  // xlCalculating
                    2 => "done",     // xlDone
                    3 => "pending",  // xlPending
                    _ => "unknown"
                };
            }
            catch
            {
                calcState = "done"; // Fallback to done if not available
            }

            return new CalculationModeResult
            {
                Success = true,
                Mode = mode,
                ModeValue = modeValue,
                CalculationState = calcState,
                IsPending = calcState == "pending",
                Message = $"Calculation mode is {mode}"
            };
        });
    }

    /// <summary>
    /// Sets the calculation mode (automatic, manual, semi-automatic).
    /// </summary>
    public OperationResult SetMode(IExcelBatch batch, CalculationMode mode)
    {
        return batch.Execute((ctx, ct) =>
        {
            int newValue = (int)mode;
            string newMode = mode switch
            {
                CalculationMode.Automatic => "automatic",
                CalculationMode.Manual => "manual",
                CalculationMode.SemiAutomatic => "semi-automatic",
                _ => "unknown"
            };

            try
            {
                ctx.App.Calculation = newValue;
            }
            catch (Exception ex)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to set calculation mode to {newMode}: {ex.Message}"
                };
            }

            return new OperationResult
            {
                Success = true,
                Message = $"Calculation mode set to {newMode}"
            };
        });
    }

    /// <summary>
    /// Triggers calculation for the specified scope (workbook, sheet, or range).
    /// </summary>
    public OperationResult Calculate(IExcelBatch batch, CalculationScope scope, string? sheetName = null, string? rangeAddress = null)
    {
        // Validate parameters
        if (scope == CalculationScope.Sheet && string.IsNullOrWhiteSpace(sheetName))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = "sheetName is required for Sheet scope calculation"
            };
        }

        if (scope == CalculationScope.Range && (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeAddress)))
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = "Both sheetName and rangeAddress are required for Range scope calculation"
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            try
            {
                switch (scope)
                {
                    case CalculationScope.Workbook:
                        ctx.App.Calculate();
                        return new OperationResult
                        {
                            Success = true,
                            Message = "Calculation complete for all workbooks"
                        };

                    case CalculationScope.Sheet:
                        dynamic? worksheet = null;
                        try
                        {
                            worksheet = ctx.Book.Worksheets[sheetName];
                            worksheet.Calculate();
                            return new OperationResult
                            {
                                Success = true,
                                Message = $"Calculation complete for sheet '{sheetName}'"
                            };
                        }
                        finally
                        {
                            if (worksheet != null) ComUtilities.Release(ref worksheet);
                        }

                    case CalculationScope.Range:
                        dynamic? ws = null;
                        dynamic? rng = null;
                        try
                        {
                            ws = ctx.Book.Worksheets[sheetName];
                            rng = ws.Range[rangeAddress];
                            rng.Calculate();
                            return new OperationResult
                            {
                                Success = true,
                                Message = $"Calculation complete for range '{rangeAddress}' on sheet '{sheetName}'"
                            };
                        }
                        finally
                        {
                            if (rng != null) ComUtilities.Release(ref rng);
                            if (ws != null) ComUtilities.Release(ref ws);
                        }

                    default:
                        return new OperationResult
                        {
                            Success = false,
                            ErrorMessage = $"Unknown calculation scope: {scope}"
                        };
                }
            }
            catch (Exception ex)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Calculation failed: {ex.Message}"
                };
            }
        });
    }
}


