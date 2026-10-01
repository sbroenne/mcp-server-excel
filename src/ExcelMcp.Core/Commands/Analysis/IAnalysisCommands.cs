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
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing the what-if model</param>
    /// <param name="formulaCell">Cell containing the formula whose result should reach the goal</param>
    /// <param name="goal">Numeric target for the formula result</param>
    /// <param name="changingCell">Single input cell Excel may adjust</param>
    [ServiceAction("goal-seek")]
    [MacCapability(
        MacCapabilityTier.Native,
        MacImplementationStatus.Implemented,
        true,
        Evidence = "PR #914 commit 6a7aaec4 passed prompt-free real-Excel Goal Seek through CLI and MCP.",
        ExcelApiVersion = "Excel for Mac 16.113.1; Apple Events/JXA.")]
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
    [MacCapability(
        MacCapabilityTier.Native,
        MacImplementationStatus.Partial,
        false,
        Evidence = "Excel 16.113.1 dictionary exposes scenario metadata and get values; portable routing is not real-Excel evidence.",
        ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary.",
        Blocker = "native API presence is not runtime parity; exact returned metadata must pass prompt-free CLI and MCP Excel tests")]
    ScenarioListResult ListScenarios(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Creates a worksheet scenario from a range of changing cells and one value per cell.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing the model</param>
    /// <param name="scenarioName">Name of the worksheet scenario</param>
    /// <param name="changingCells">Range of input cells whose values the scenario stores</param>
    /// <param name="values">One value per changing cell, in range order</param>
    /// <param name="comment">Optional scenario description</param>
    /// <param name="locked">Prevent scenario editing when worksheet protection is enabled</param>
    /// <param name="hidden">Hide the scenario when worksheet protection is enabled</param>
    [ServiceAction("create-scenario")]
    [MacCapability(
        MacCapabilityTier.Unsupported,
        MacImplementationStatus.Blocked,
        false,
        Evidence = "Excel for Mac 16.113.1 Apple Events exposes scenario elements but no scenario creation command, and Office.js exposes no Scenario API.",
        ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary; Office.js API review.",
        Blocker = "no supported local macOS API can create a scenario; ExcelMcp does not ship a VBA helper")]
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
    [MacCapability(
        MacCapabilityTier.Native,
        MacImplementationStatus.Partial,
        false,
        Evidence = "Excel 16.113.1 dictionary exposes change scenario; portable routing is not real-Excel evidence.",
        ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary.",
        Blocker = "native API presence is not runtime parity; exact changed cells and values must pass prompt-free CLI and MCP Excel tests")]
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
    [MacCapability(
        MacCapabilityTier.Unsupported,
        MacImplementationStatus.Blocked,
        false,
        Evidence = "Excel for Mac 16.113.1 Apple Events exposes scenario metadata and change/delete/summary commands but no show-scenario command, and Office.js exposes no Scenario API.",
        ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary; Office.js API review.",
        Blocker = "no supported local macOS API can apply stored scenario values; ExcelMcp does not ship a VBA helper")]
    OperationResult ShowScenario(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string scenarioName);

    /// <summary>
    /// Deletes a worksheet scenario.
    /// </summary>
    [ServiceAction("delete-scenario")]
    [MacCapability(
        MacCapabilityTier.Native,
        MacImplementationStatus.Partial,
        false,
        Evidence = "Excel 16.113.1 dictionary exposes scenario elements and generic delete; portable routing is not real-Excel evidence.",
        ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary.",
        Blocker = "native API presence is not runtime parity; exact-name deletion and worksheet effects must pass prompt-free CLI and MCP Excel tests")]
    OperationResult DeleteScenario(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string scenarioName);

    /// <summary>
    /// Creates a standard worksheet summary or PivotTable summary for all scenarios on a worksheet.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing the scenarios</param>
    /// <param name="reportType">Scenario report type: Summary or PivotTable</param>
    /// <param name="resultCells">Formula result cells to include in the scenario report</param>
    [ServiceAction("create-scenario-summary")]
    [MacCapability(
        MacCapabilityTier.Native,
        MacImplementationStatus.Partial,
        false,
        Evidence = "Excel 16.113.1 dictionary exposes create summary for scenarios with standard and PivotTable report types.",
        ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary.",
        Blocker = "native API presence is not runtime parity; new-sheet identity and both summary types must pass prompt-free CLI and MCP Excel tests")]
    ScenarioSummaryResult CreateScenarioSummary(
        IExcelBatch batch,
        string sheetName,
        [FromString("reportType")] ScenarioSummaryType reportType = ScenarioSummaryType.Summary,
        string? resultCells = null);

    /// <summary>
    /// Creates a one- or two-variable Excel data table from a prepared formula and input-value range.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet containing the prepared what-if model</param>
    /// <param name="tableRange">Prepared sensitivity table range, including formulas and trial input values</param>
    /// <param name="rowInputCell">Model input cell to substitute values from the table's row; supply at least one input cell</param>
    /// <param name="columnInputCell">Model input cell to substitute values from the table's column; supply at least one input cell</param>
    [ServiceAction("create-data-table")]
    [MacCapability(
        MacCapabilityTier.Native,
        MacImplementationStatus.Implemented,
        true,
        Evidence = "PR #914 commit 6a7aaec4 passed a prompt-free literal Data Table through CLI and MCP.",
        ExcelApiVersion = "Excel for Mac 16.113.1; Apple Events/JXA.")]
    OperationResult CreateDataTable(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string tableRange,
        string? rowInputCell = null,
        string? columnInputCell = null);
}
