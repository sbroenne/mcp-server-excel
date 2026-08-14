using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Analysis;

public sealed partial class AnalysisCommands
{
    /// <inheritdoc />
    public ScenarioListResult ListScenarios(IExcelBatch batch, string sheetName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Scenarios? scenarios = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
                scenarios = (Excel.Scenarios)sheet.Scenarios(Type.Missing);

                var result = new ScenarioListResult
                {
                    Success = true,
                    SheetName = sheetName,
                    Message = $"Found {scenarios.Count} scenario(s) on '{sheetName}'."
                };

                for (var index = 1; index <= scenarios.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Scenario? scenario = null;
                    Excel.Range? changingCells = null;
                    try
                    {
                        scenario = scenarios.Item(index);
                        changingCells = scenario.ChangingCells;
                        result.Scenarios.Add(new ScenarioInfo
                        {
                            Name = scenario.Name,
                            ChangingCells = changingCells.Address[true, true],
                            Values = ConvertScenarioValues(scenario.get_Values(Type.Missing)),
                            Comment = scenario.Comment,
                            Locked = scenario.Locked,
                            Hidden = scenario.Hidden
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref changingCells);
                        ComUtilities.Release(ref scenario);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref scenarios);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult CreateScenario(
        IExcelBatch batch,
        string sheetName,
        string scenarioName,
        string changingCells,
        List<object?> values,
        string? comment = null,
        bool locked = true,
        bool hidden = false)
    {
        ValidateScenarioArguments(sheetName, scenarioName, changingCells, values);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? changingRange = null;
            Excel.Scenarios? scenarios = null;
            Excel.Scenario? scenario = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
                changingRange = sheet.Range[changingCells];
                ValidateScenarioValueCount(changingRange, values);
                scenarios = (Excel.Scenarios)sheet.Scenarios(Type.Missing);
                scenario = scenarios.Add(
                    scenarioName,
                    changingRange,
                    ConvertScenarioInputValues(values),
                    (object?)comment ?? Type.Missing,
                    locked,
                    hidden);

                return new OperationResult
                {
                    Success = true,
                    Message = $"Scenario '{scenarioName}' created on '{sheetName}'."
                };
            }
            finally
            {
                ComUtilities.Release(ref scenario);
                ComUtilities.Release(ref scenarios);
                ComUtilities.Release(ref changingRange);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult UpdateScenario(
        IExcelBatch batch,
        string sheetName,
        string scenarioName,
        string changingCells,
        List<object?> values)
    {
        ValidateScenarioArguments(sheetName, scenarioName, changingCells, values);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? changingRange = null;
            Excel.Scenarios? scenarios = null;
            Excel.Scenario? scenario = null;
            object? changeResult = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
                changingRange = sheet.Range[changingCells];
                ValidateScenarioValueCount(changingRange, values);
                scenarios = (Excel.Scenarios)sheet.Scenarios(Type.Missing);
                scenario = scenarios.Item(scenarioName);
                changeResult = scenario.ChangeScenario(changingRange, ConvertScenarioInputValues(values));

                return new OperationResult
                {
                    Success = true,
                    Message = $"Scenario '{scenarioName}' updated on '{sheetName}'."
                };
            }
            finally
            {
                ComUtilities.Release(ref changeResult);
                ComUtilities.Release(ref scenario);
                ComUtilities.Release(ref scenarios);
                ComUtilities.Release(ref changingRange);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult ShowScenario(IExcelBatch batch, string sheetName, string scenarioName)
    {
        return ExecuteScenarioAction(batch, sheetName, scenarioName, "shown", scenario => scenario.Show());
    }

    /// <inheritdoc />
    public OperationResult DeleteScenario(IExcelBatch batch, string sheetName, string scenarioName)
    {
        return ExecuteScenarioAction(batch, sheetName, scenarioName, "deleted", scenario => scenario.Delete());
    }

    /// <inheritdoc />
    public ScenarioSummaryResult CreateScenarioSummary(
        IExcelBatch batch,
        string sheetName,
        ScenarioSummaryType reportType = ScenarioSummaryType.Summary,
        string? resultCells = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? resultRange = null;
            Excel.Scenarios? scenarios = null;
            object? summaryResult = null;
            try
            {
                var existingSheetNames = GetWorksheetNames(ctx.Book);
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
                if (!string.IsNullOrWhiteSpace(resultCells))
                {
                    resultRange = sheet.Range[resultCells];
                }

                scenarios = (Excel.Scenarios)sheet.Scenarios(Type.Missing);
                var excelReportType = reportType switch
                {
                    ScenarioSummaryType.Summary => Excel.XlSummaryReportType.xlStandardSummary,
                    ScenarioSummaryType.PivotTable => Excel.XlSummaryReportType.xlSummaryPivotTable,
                    _ => throw new ArgumentOutOfRangeException(nameof(reportType), reportType, "Unknown scenario summary type.")
                };
                summaryResult = scenarios.CreateSummary(excelReportType, (object?)resultRange ?? Type.Missing);

                return new ScenarioSummaryResult
                {
                    Success = true,
                    ReportSheetName = FindNewWorksheetName(ctx.Book, existingSheetNames),
                    ReportType = reportType == ScenarioSummaryType.Summary ? "summary" : "pivot-table",
                    Message = $"Scenario {reportType.ToString().ToLowerInvariant()} report created."
                };
            }
            finally
            {
                ComUtilities.Release(ref summaryResult);
                ComUtilities.Release(ref scenarios);
                ComUtilities.Release(ref resultRange);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static OperationResult ExecuteScenarioAction(
        IExcelBatch batch,
        string sheetName,
        string scenarioName,
        string pastTense,
        Func<Excel.Scenario, object> action)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(scenarioName);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Scenarios? scenarios = null;
            Excel.Scenario? scenario = null;
            object? actionResult = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
                scenarios = (Excel.Scenarios)sheet.Scenarios(Type.Missing);
                scenario = scenarios.Item(scenarioName);
                actionResult = action(scenario);

                return new OperationResult
                {
                    Success = true,
                    Message = $"Scenario '{scenarioName}' {pastTense} on '{sheetName}'."
                };
            }
            finally
            {
                ComUtilities.Release(ref actionResult);
                ComUtilities.Release(ref scenario);
                ComUtilities.Release(ref scenarios);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void ValidateScenarioArguments(
        string sheetName,
        string scenarioName,
        string changingCells,
        List<object?> values)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(scenarioName);
        ArgumentException.ThrowIfNullOrWhiteSpace(changingCells);
        ArgumentNullException.ThrowIfNull(values);
        if (values.Count == 0)
        {
            throw new ArgumentException("At least one scenario value is required.", nameof(values));
        }
    }

    private static void ValidateScenarioValueCount(Excel.Range changingRange, List<object?> values)
    {
        var changingCellCount = Convert.ToInt32(changingRange.CountLarge);
        if (changingCellCount != values.Count)
        {
            throw new ArgumentException(
                $"Scenario values count ({values.Count}) must match changing cells count ({changingCellCount}).",
                nameof(values));
        }
    }

    private static object[] ConvertScenarioInputValues(List<object?> values)
    {
        return values.Select(RangeHelpers.ConvertToCellValue).ToArray();
    }

    private static List<object?> ConvertScenarioValues(object? values)
    {
        if (values is not Array array)
        {
            return [values];
        }

        var result = new List<object?>(array.Length);
        foreach (var value in array)
        {
            result.Add(value);
        }

        return result;
    }

    private static HashSet<string> GetWorksheetNames(Excel.Workbook workbook)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            for (var index = 1; index <= sheets.Count; index++)
            {
                Excel.Worksheet? sheet = null;
                try
                {
                    sheet = (Excel.Worksheet)sheets[index];
                    names.Add(sheet.Name);
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }

        return names;
    }

    private static string FindNewWorksheetName(Excel.Workbook workbook, HashSet<string> existingNames)
    {
        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            for (var index = 1; index <= sheets.Count; index++)
            {
                Excel.Worksheet? sheet = null;
                try
                {
                    sheet = (Excel.Worksheet)sheets[index];
                    if (!existingNames.Contains(sheet.Name))
                    {
                        return sheet.Name;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }

        throw new InvalidOperationException("Excel did not create a scenario summary worksheet.");
    }
}
