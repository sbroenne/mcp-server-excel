using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Analysis;

/// <summary>
/// Excel what-if analysis with Goal Seek, scenarios, scenario reports, and one- or two-variable data tables.
/// Solver is excluded because it is an optional VBA add-in that must be enabled by the user and is not part of the Excel PIA.
/// </summary>
[ServiceCategory("analysis", "Analysis")]
[McpTool("analysis", Title = "What-If Analysis", Destructive = true, Category = "analysis",
    Description = "Run Excel what-if analysis using the native Excel COM object model. GOAL SEEK adjusts one input cell until a formula reaches a numeric goal. SCENARIOS create, list, update, show, delete, and summarize named input sets on a worksheet. DATA TABLES create one- or two-variable sensitivity tables from a prepared worksheet model. Solver is not exposed because Microsoft implements it as an optional VBA add-in that must be manually enabled and referenced, not as a reliable Excel PIA API.")]
public interface IAnalysisCommands
{
    /// <summary>
    /// Adjusts one changing cell until a formula cell reaches the requested numeric goal.
    /// </summary>
    [ServiceAction("goal-seek")]
    GoalSeekResult GoalSeek(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string formulaCell,
        [RequiredParameter] double? goal,
        [RequiredParameter] string changingCell);

    /// <summary>
    /// Lists the scenarios defined on a worksheet, including changing cells, values, and protection metadata.
    /// </summary>
    [ServiceAction("list-scenarios")]
    ScenarioListResult ListScenarios(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Creates a worksheet scenario from a range of changing cells and one value per cell.
    /// </summary>
    [ServiceAction("create-scenario")]
    OperationResult CreateScenario(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string scenarioName,
        [RequiredParameter] string changingCells,
        [RequiredParameter] List<object?> values,
        string? comment = null,
        bool locked = true,
        bool hidden = false);

    /// <summary>
    /// Replaces the changing cells and values of an existing worksheet scenario.
    /// </summary>
    [ServiceAction("update-scenario")]
    OperationResult UpdateScenario(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string scenarioName,
        [RequiredParameter] string changingCells,
        [RequiredParameter] List<object?> values);

    /// <summary>
    /// Applies a scenario's stored values to its changing cells.
    /// </summary>
    [ServiceAction("show-scenario")]
    OperationResult ShowScenario(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string scenarioName);

    /// <summary>
    /// Deletes a worksheet scenario.
    /// </summary>
    [ServiceAction("delete-scenario")]
    OperationResult DeleteScenario(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string scenarioName);

    /// <summary>
    /// Creates a standard worksheet summary or PivotTable summary for all scenarios on a worksheet.
    /// </summary>
    [ServiceAction("create-scenario-summary")]
    ScenarioSummaryResult CreateScenarioSummary(
        IExcelBatch batch,
        string sheetName,
        [FromString("reportType")] ScenarioSummaryType reportType = ScenarioSummaryType.Summary,
        string? resultCells = null);

    /// <summary>
    /// Creates a one- or two-variable Excel data table from a prepared formula and input-value range.
    /// </summary>
    [ServiceAction("create-data-table")]
    OperationResult CreateDataTable(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string tableRange,
        string? rowInputCell = null,
        string? columnInputCell = null);
}
