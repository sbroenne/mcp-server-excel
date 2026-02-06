using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel calculation mode control.
/// </summary>
[McpServerToolType]
public static partial class ExcelCalculationModeTool
{
    /// <summary>
    /// Set or get Excel calculation mode and explicitly recalculate formulas. Use this tool whenever a task mentions calculation mode, manual/automatic calculation, or explicit recalculation. Do NOT use excel_range or excel_worksheet for these actions.
    ///
    /// MODES:
    /// - 'automatic': Recalculates on every change (default)
    /// - 'manual': Only recalculates when explicitly requested
    /// - 'semi-automatic': Auto except data tables (recalc-intensive)
    ///
    /// WORKFLOW for batch operations:
    /// 1. set-mode(mode='manual') - Disable auto-recalc
    /// 2. Perform data operations (excel_range set-values, etc.)
    /// 3. calculate(scope='workbook') - Recalculate once
    /// 4. set-mode(mode='automatic') - Restore default
    ///
    /// SCOPES for calculate action:
    /// - 'workbook': Recalculate all formulas
    /// - 'sheet': Recalculate single sheet (requires sheetName)
    /// - 'range': Recalculate specific range (requires sheetName + rangeAddress)
    /// </summary>
    /// <param name="action">Action to perform: get-mode, set-mode, calculate</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action</param>
    /// <param name="mode">Calculation mode for set-mode action: 'automatic', 'manual', 'semi-automatic'</param>
    /// <param name="scope">Calculation scope for calculate action: 'workbook', 'sheet', 'range'</param>
    /// <param name="sheetName">Sheet name (required for sheet/range scope)</param>
    /// <param name="rangeAddress">Range address (required for range scope, e.g., 'A1:C10')</param>
    [McpServerTool(Name = "excel_calculation_mode", Title = "Excel Calculation Mode Control", Destructive = false)]
    [McpMeta("category", "settings")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelCalculationMode(
        CalculationAction action,
        string sessionId,
        [DefaultValue(null)] string? mode,
        [DefaultValue(null)] string? scope,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? rangeAddress)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_calculation_mode",
            ServiceRegistry.Calculation.ToActionString(action),
            () => ServiceRegistry.Calculation.RouteAction(
                action,
                sessionId,
                ExcelToolsBase.ForwardToServiceFunc,
                mode: mode,
                scope: scope,
                sheetName: sheetName,
                rangeAddress: rangeAddress));
    }
}
