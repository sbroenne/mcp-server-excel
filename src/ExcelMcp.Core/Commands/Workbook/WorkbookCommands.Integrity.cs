using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    private const int MaximumIntegrityFindings = 10_000;
    private const int ExcelNoMatchingCellsHResult = unchecked((int)0x800A03EC);

    /// <inheritdoc />
    public WorkbookIntegrityResult ValidateIntegrity(
        IExcelBatch batch,
        List<WorkbookIntegrityCheck>? checks = null,
        List<string>? worksheetNames = null,
        List<WorkbookControlTotalExpectation>? controlTotals = null,
        int maxFindings = 500)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ValidateIntegrityInputs(checks, worksheetNames, controlTotals, maxFindings);

        var selectedChecks = ResolveChecks(checks, controlTotals);
        var selectedWorksheetNames = worksheetNames is null
            ? null
            : new HashSet<string>(worksheetNames, StringComparer.OrdinalIgnoreCase);
        var result = new WorkbookIntegrityResult
        {
            FilePath = batch.WorkbookPath,
            CheckedChecks = selectedChecks
        };

        return batch.Execute((context, cancellationToken) =>
        {
            bool wasSaved = context.Book.Saved;
            try
            {
                var findings = new IntegrityFindingCollector(result, maxFindings);
                ReadCalculationState(context.App, result, selectedChecks, findings);

                if (selectedChecks.Contains(WorkbookIntegrityCheck.FormulaErrors) ||
                    selectedChecks.Contains(WorkbookIntegrityCheck.Tables))
                {
                    ValidateWorksheets(
                        context.Book,
                        selectedChecks,
                        selectedWorksheetNames,
                        result,
                        findings,
                        cancellationToken);
                }

                if (selectedChecks.Contains(WorkbookIntegrityCheck.ExternalLinks))
                {
                    ValidateExternalLinks(context.Book, findings, cancellationToken);
                }

                if (selectedChecks.Contains(WorkbookIntegrityCheck.ControlTotals))
                {
                    ValidateControlTotals(
                        context.Book,
                        controlTotals!,
                        result,
                        findings,
                        cancellationToken);
                }

                findings.Complete();
                result.Success = true;
                return result;
            }
            finally
            {
                if (context.Book.Saved != wasSaved)
                {
                    context.Book.Saved = wasSaved;
                }
            }
        });
    }

    private static void ValidateIntegrityInputs(
        List<WorkbookIntegrityCheck>? checks,
        List<string>? worksheetNames,
        List<WorkbookControlTotalExpectation>? controlTotals,
        int maxFindings)
    {
        if (maxFindings is < 1 or > MaximumIntegrityFindings)
        {
            throw new ArgumentOutOfRangeException(
                nameof(maxFindings),
                maxFindings,
                $"Maximum findings must be between 1 and {MaximumIntegrityFindings}.");
        }

        if (checks is { Count: 0 })
        {
            throw new ArgumentException("At least one integrity check is required when checks are supplied.", nameof(checks));
        }

        if (checks is not null)
        {
            foreach (var check in checks)
            {
                if (!Enum.IsDefined(check))
                {
                    throw new ArgumentOutOfRangeException(nameof(checks), check, "Unknown workbook integrity check.");
                }
            }

            if (checks.Contains(WorkbookIntegrityCheck.ControlTotals) &&
                controlTotals is not { Count: > 0 })
            {
                throw new ArgumentException(
                    "Control-total expectations are required when the control-totals check is selected.",
                    nameof(controlTotals));
            }

            if (!checks.Contains(WorkbookIntegrityCheck.ControlTotals) &&
                controlTotals is { Count: > 0 })
            {
                throw new ArgumentException(
                    "Control-total expectations require the control-totals check when checks are supplied.",
                    nameof(controlTotals));
            }
        }

        if (worksheetNames is { Count: 0 })
        {
            throw new ArgumentException(
                "At least one worksheet name is required when worksheetNames is supplied.",
                nameof(worksheetNames));
        }

        if (worksheetNames is not null)
        {
            if (checks is not null &&
                !checks.Contains(WorkbookIntegrityCheck.FormulaErrors) &&
                !checks.Contains(WorkbookIntegrityCheck.Tables))
            {
                throw new ArgumentException(
                    "Worksheet names apply only to formula-errors and tables checks.",
                    nameof(worksheetNames));
            }

            foreach (var worksheetName in worksheetNames)
            {
                if (string.IsNullOrWhiteSpace(worksheetName))
                {
                    throw new ArgumentException("Worksheet names cannot be empty.", nameof(worksheetNames));
                }
            }
        }

        if (controlTotals is null)
        {
            return;
        }

        if (controlTotals.Count == 0)
        {
            throw new ArgumentException(
                "At least one control-total expectation is required when controlTotals is supplied.",
                nameof(controlTotals));
        }

        for (int index = 0; index < controlTotals.Count; index++)
        {
            var expectation = controlTotals[index]
                ?? throw new ArgumentException($"Control total at index {index} cannot be null.", nameof(controlTotals));
            if (string.IsNullOrWhiteSpace(expectation.SheetName))
            {
                throw new ArgumentException(
                    $"Control total at index {index} requires a worksheet name.",
                    nameof(controlTotals));
            }

            if (string.IsNullOrWhiteSpace(expectation.CellAddress))
            {
                throw new ArgumentException(
                    $"Control total at index {index} requires a cell address.",
                    nameof(controlTotals));
            }

            if (!expectation.ExpectedValue.HasValue ||
                !double.IsFinite(expectation.ExpectedValue.Value))
            {
                throw new ArgumentException(
                    $"Control total at index {index} requires a finite expected value.",
                    nameof(controlTotals));
            }

            if (!double.IsFinite(expectation.Tolerance) || expectation.Tolerance < 0)
            {
                throw new ArgumentException(
                    $"Control total at index {index} requires a finite, non-negative tolerance.",
                    nameof(controlTotals));
            }
        }
    }

    private static List<WorkbookIntegrityCheck> ResolveChecks(
        List<WorkbookIntegrityCheck>? checks,
        List<WorkbookControlTotalExpectation>? controlTotals)
    {
        if (checks is not null)
        {
            return checks.Distinct().ToList();
        }

        var result = new List<WorkbookIntegrityCheck>
        {
            WorkbookIntegrityCheck.FormulaErrors,
            WorkbookIntegrityCheck.ExternalLinks,
            WorkbookIntegrityCheck.Tables
        };
        if (controlTotals is { Count: > 0 })
        {
            result.Add(WorkbookIntegrityCheck.ControlTotals);
        }

        return result;
    }

    private static void ReadCalculationState(
        Excel.Application application,
        WorkbookIntegrityResult result,
        List<WorkbookIntegrityCheck> checks,
        IntegrityFindingCollector findings)
    {
        var calculationMode = application.Calculation;
        var calculationState = application.CalculationState;
        result.CalculationMode = calculationMode switch
        {
            Excel.XlCalculation.xlCalculationAutomatic => "automatic",
            Excel.XlCalculation.xlCalculationManual => "manual",
            Excel.XlCalculation.xlCalculationSemiautomatic => "semi-automatic",
            _ => Convert.ToInt32(calculationMode, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture)
        };
        result.CalculationState = calculationState switch
        {
            Excel.XlCalculationState.xlDone => "done",
            Excel.XlCalculationState.xlCalculating => "calculating",
            Excel.XlCalculationState.xlPending => "pending",
            _ => Convert.ToInt32(calculationState, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture)
        };

        bool readsCalculatedValues =
            checks.Contains(WorkbookIntegrityCheck.FormulaErrors) ||
            checks.Contains(WorkbookIntegrityCheck.ControlTotals);
        if (!readsCalculatedValues)
        {
            return;
        }

        if (calculationMode == Excel.XlCalculation.xlCalculationManual)
        {
            findings.Add(new WorkbookIntegrityFinding
            {
                Code = "manual-calculation",
                Severity = WorkbookIntegritySeverity.Warning,
                Category = WorkbookIntegrityCategory.CalculationState,
                Reliability = WorkbookIntegrityReliability.Heuristic,
                Message = "Workbook calculation mode is manual, so cached formula and control-total values may be stale.",
                SuggestedRemediation = "Calculate the workbook, then run integrity validation again."
            });
        }

        if (calculationState != Excel.XlCalculationState.xlDone)
        {
            findings.Add(new WorkbookIntegrityFinding
            {
                Code = "calculation-incomplete",
                Severity = WorkbookIntegritySeverity.Warning,
                Category = WorkbookIntegrityCategory.CalculationState,
                Reliability = WorkbookIntegrityReliability.Heuristic,
                Message = $"Workbook calculation state is {result.CalculationState}, so calculated values may be incomplete.",
                SuggestedRemediation = "Wait for calculation to finish, then run integrity validation again."
            });
        }
    }

    private static void ValidateWorksheets(
        Excel.Workbook workbook,
        List<WorkbookIntegrityCheck> checks,
        HashSet<string>? selectedWorksheetNames,
        WorkbookIntegrityResult result,
        IntegrityFindingCollector findings,
        CancellationToken cancellationToken)
    {
        Excel.Sheets? worksheets = null;
        var unmatchedNames = selectedWorksheetNames is null
            ? null
            : new HashSet<string>(selectedWorksheetNames, StringComparer.OrdinalIgnoreCase);
        try
        {
            worksheets = workbook.Worksheets;
            int worksheetCount = worksheets.Count;
            for (int index = 1; index <= worksheetCount; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.Worksheet? worksheet = null;
                try
                {
                    worksheet = (Excel.Worksheet)worksheets.Item[index];
                    string worksheetName = worksheet.Name;
                    if (selectedWorksheetNames is not null &&
                        !selectedWorksheetNames.Contains(worksheetName))
                    {
                        continue;
                    }

                    unmatchedNames?.Remove(worksheetName);
                    result.CheckedWorksheets.Add(worksheetName);
                    if (checks.Contains(WorkbookIntegrityCheck.FormulaErrors))
                    {
                        ValidateFormulaErrors(worksheet, worksheetName, findings, cancellationToken);
                    }

                    if (checks.Contains(WorkbookIntegrityCheck.Tables))
                    {
                        ValidateTables(worksheet, worksheetName, findings, cancellationToken);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref worksheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }

        if (unmatchedNames is { Count: > 0 })
        {
            throw new InvalidOperationException(
                $"Worksheet(s) not found: {string.Join(", ", unmatchedNames.Order(StringComparer.OrdinalIgnoreCase))}.");
        }
    }

    private static void ValidateFormulaErrors(
        Excel.Worksheet worksheet,
        string worksheetName,
        IntegrityFindingCollector findings,
        CancellationToken cancellationToken)
    {
        Excel.Range? usedRange = null;
        Excel.Range? errorRange = null;
        Excel.Areas? areas = null;
        var reportedErrorCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            usedRange = worksheet.UsedRange;
            try
            {
                errorRange = usedRange.SpecialCells(
                    Excel.XlCellType.xlCellTypeFormulas,
                    Excel.XlSpecialCellsValue.xlErrors);
            }
            catch (COMException exception) when (exception.HResult == ExcelNoMatchingCellsHResult)
            {
                ScanFormulaErrorsInRange(
                    usedRange,
                    worksheetName,
                    findings,
                    reportedErrorCells,
                    cancellationToken);
                ValidateBrokenReferenceTokens(
                    usedRange,
                    worksheetName,
                    findings,
                    reportedErrorCells,
                    cancellationToken);
                return;
            }

            areas = errorRange.Areas;
            int areaCount = areas.Count;
            for (int areaIndex = 1; areaIndex <= areaCount; areaIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.Range? area = null;
                try
                {
                    area = areas.Item[areaIndex];
                    var (rowCount, columnCount) = GetRangeDimensions(area);
                    int startRow = area.Row;
                    int startColumn = area.Column;
                    object values = area.Value2;
                    object formulas = area.Formula2;

                    for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
                        {
                            if ((columnOffset & 255) == 0)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                            }

                            object? value = GetMatrixValue(values, rowOffset, columnOffset);
                            string cellAddress = GetCellAddress(
                                startRow + rowOffset,
                                startColumn + columnOffset);
                            if (!ExcelErrorMapper.TryGet(value, out int errorCode, out var error))
                            {
                                if (value is int unknownErrorCode)
                                {
                                    reportedErrorCells.Add(cellAddress);
                                    findings.Add(new WorkbookIntegrityFinding
                                    {
                                        Code = "formula-error-unknown",
                                        Severity = WorkbookIntegritySeverity.Error,
                                        Category = WorkbookIntegrityCategory.FormulaError,
                                        Reliability = WorkbookIntegrityReliability.Deterministic,
                                        Message = ExcelErrorMapper.GetMessage(unknownErrorCode),
                                        SuggestedRemediation = "Open the affected cell in Excel and review its formula and precedents.",
                                        SheetName = worksheetName,
                                        CellAddress = cellAddress,
                                        Formula = Convert.ToString(
                                            GetMatrixValue(formulas, rowOffset, columnOffset),
                                            CultureInfo.InvariantCulture),
                                        ErrorName = "#ERROR!",
                                        ErrorCode = unknownErrorCode,
                                        ActualValue = unknownErrorCode
                                    });
                                }

                                continue;
                            }

                            string? formula = Convert.ToString(
                                GetMatrixValue(formulas, rowOffset, columnOffset),
                                CultureInfo.InvariantCulture);
                            var category = error.Name == "#REF!"
                                ? WorkbookIntegrityCategory.BrokenReference
                                : WorkbookIntegrityCategory.FormulaError;
                            reportedErrorCells.Add(cellAddress);
                            findings.Add(new WorkbookIntegrityFinding
                            {
                                Code = category == WorkbookIntegrityCategory.BrokenReference
                                    ? "broken-formula-reference"
                                    : "formula-error",
                                Severity = IsTransientFormulaState(error.Name)
                                    ? WorkbookIntegritySeverity.Warning
                                    : WorkbookIntegritySeverity.Error,
                                Category = category,
                                Reliability = WorkbookIntegrityReliability.Deterministic,
                                Message = $"{error.Name} - {error.Description}",
                                SuggestedRemediation = error.Suggestion,
                                SheetName = worksheetName,
                                CellAddress = cellAddress,
                                Formula = formula,
                                ErrorName = error.Name,
                                ErrorCode = errorCode,
                                ActualValue = error.Name
                            });
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref area);
                }
            }

            ValidateBrokenReferenceTokens(
                usedRange,
                worksheetName,
                findings,
                reportedErrorCells,
                cancellationToken);
        }
        finally
        {
            ComUtilities.Release(ref areas);
            ComUtilities.Release(ref errorRange);
            ComUtilities.Release(ref usedRange);
        }
    }

    private static bool IsTransientFormulaState(string errorName) =>
        errorName is "#GETTING_DATA" or "#CONNECT!" or "#BUSY!";

    private static void ScanFormulaErrorsInRange(
        Excel.Range range,
        string worksheetName,
        IntegrityFindingCollector findings,
        HashSet<string> reportedErrorCells,
        CancellationToken cancellationToken)
    {
        var (rowCount, columnCount) = GetRangeDimensions(range);
        int startRow = range.Row;
        int startColumn = range.Column;
        object values = range.Value2;
        object formulas = range.Formula2;

        for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
            {
                if ((columnOffset & 255) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }

                object? value = GetMatrixValue(values, rowOffset, columnOffset);
                if (value is not int possibleErrorCode)
                {
                    continue;
                }

                string cellAddress = GetCellAddress(
                    startRow + rowOffset,
                    startColumn + columnOffset);
                string? formula = Convert.ToString(
                    GetMatrixValue(formulas, rowOffset, columnOffset),
                    CultureInfo.InvariantCulture);
                if (formula?.StartsWith('=') != true)
                {
                    continue;
                }

                reportedErrorCells.Add(cellAddress);
                if (!ExcelErrorMapper.TryGet(possibleErrorCode, out var error))
                {
                    findings.Add(new WorkbookIntegrityFinding
                    {
                        Code = "formula-error-unknown",
                        Severity = WorkbookIntegritySeverity.Error,
                        Category = WorkbookIntegrityCategory.FormulaError,
                        Reliability = WorkbookIntegrityReliability.Deterministic,
                        Message = ExcelErrorMapper.GetMessage(possibleErrorCode),
                        SuggestedRemediation = "Open the affected cell in Excel and review its formula and precedents.",
                        SheetName = worksheetName,
                        CellAddress = cellAddress,
                        Formula = formula,
                        ErrorName = "#ERROR!",
                        ErrorCode = possibleErrorCode,
                        ActualValue = possibleErrorCode
                    });
                    continue;
                }

                var category = error.Name == "#REF!"
                    ? WorkbookIntegrityCategory.BrokenReference
                    : WorkbookIntegrityCategory.FormulaError;
                findings.Add(new WorkbookIntegrityFinding
                {
                    Code = category == WorkbookIntegrityCategory.BrokenReference
                        ? "broken-formula-reference"
                        : "formula-error",
                    Severity = IsTransientFormulaState(error.Name)
                        ? WorkbookIntegritySeverity.Warning
                        : WorkbookIntegritySeverity.Error,
                    Category = category,
                    Reliability = WorkbookIntegrityReliability.Deterministic,
                    Message = $"{error.Name} - {error.Description}",
                    SuggestedRemediation = error.Suggestion,
                    SheetName = worksheetName,
                    CellAddress = cellAddress,
                    Formula = formula,
                    ErrorName = error.Name,
                    ErrorCode = possibleErrorCode,
                    ActualValue = error.Name
                });
            }
        }
    }

    private static void ValidateBrokenReferenceTokens(
        Excel.Range usedRange,
        string worksheetName,
        IntegrityFindingCollector findings,
        HashSet<string> reportedErrorCells,
        CancellationToken cancellationToken)
    {
        object hasFormula = usedRange.HasFormula;
        if (hasFormula is bool hasAnyFormula && !hasAnyFormula)
        {
            return;
        }

        Excel.Range? formulaCells = null;
        Excel.Areas? areas = null;
        try
        {
            try
            {
                formulaCells = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
            }
            catch (COMException exception) when (exception.HResult == ExcelNoMatchingCellsHResult)
            {
                ScanBrokenReferenceTokensInRange(
                    usedRange,
                    worksheetName,
                    findings,
                    reportedErrorCells,
                    cancellationToken);
                return;
            }

            areas = formulaCells.Areas;
            int areaCount = areas.Count;
            for (int areaIndex = 1; areaIndex <= areaCount; areaIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.Range? area = null;
                try
                {
                    area = areas.Item[areaIndex];
                    ScanBrokenReferenceTokensInRange(
                        area,
                        worksheetName,
                        findings,
                        reportedErrorCells,
                        cancellationToken);
                }
                finally
                {
                    ComUtilities.Release(ref area);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref areas);
            ComUtilities.Release(ref formulaCells);
        }
    }

    private static void ScanBrokenReferenceTokensInRange(
        Excel.Range range,
        string worksheetName,
        IntegrityFindingCollector findings,
        HashSet<string> reportedErrorCells,
        CancellationToken cancellationToken)
    {
        var (rowCount, columnCount) = GetRangeDimensions(range);
        int startRow = range.Row;
        int startColumn = range.Column;
        object formulas = range.Formula2;
        for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
            {
                string cellAddress = GetCellAddress(
                    startRow + rowOffset,
                    startColumn + columnOffset);
                if (reportedErrorCells.Contains(cellAddress))
                {
                    continue;
                }

                string? formula = Convert.ToString(
                    GetMatrixValue(formulas, rowOffset, columnOffset),
                    CultureInfo.InvariantCulture);
                if (formula?.StartsWith('=') != true ||
                    !ContainsBrokenReferenceTokenOutsideString(formula))
                {
                    continue;
                }

                findings.Add(new WorkbookIntegrityFinding
                {
                    Code = "broken-reference-token",
                    Severity = WorkbookIntegritySeverity.Error,
                    Category = WorkbookIntegrityCategory.BrokenReference,
                    Reliability = WorkbookIntegrityReliability.Deterministic,
                    Message = "Formula contains a broken #REF! reference even though error handling may hide it.",
                    SuggestedRemediation = "Replace the #REF! token with a valid cell, range, sheet, or workbook reference.",
                    SheetName = worksheetName,
                    CellAddress = cellAddress,
                    Formula = formula,
                    ErrorName = "#REF!"
                });
            }
        }
    }

    private static bool ContainsBrokenReferenceTokenOutsideString(string formula)
    {
        bool inString = false;
        for (int index = 0; index <= formula.Length - 5; index++)
        {
            if (formula[index] == '"')
            {
                if (inString && index + 1 < formula.Length && formula[index + 1] == '"')
                {
                    index++;
                    continue;
                }

                inString = !inString;
                continue;
            }

            if (!inString &&
                formula.AsSpan(index).StartsWith("#REF!".AsSpan(), StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static void ValidateExternalLinks(
        Excel.Workbook workbook,
        IntegrityFindingCollector findings,
        CancellationToken cancellationToken)
    {
        foreach (string source in GetExternalLinkSources(workbook))
        {
            cancellationToken.ThrowIfCancellationRequested();
            int status;
            try
            {
                status = Convert.ToInt32(
                    workbook.LinkInfo(source, Excel.XlLinkInfo.xlLinkInfoStatus),
                    CultureInfo.InvariantCulture);
            }
            catch (COMException exception) when (exception.HResult == ExcelNoMatchingCellsHResult)
            {
                findings.Add(new WorkbookIntegrityFinding
                {
                    Code = "external-link-status-unavailable",
                    Severity = WorkbookIntegritySeverity.Warning,
                    Category = WorkbookIntegrityCategory.ExternalLink,
                    Reliability = WorkbookIntegrityReliability.Deterministic,
                    Message = $"Excel could not determine the status of external link '{source}'.",
                    SuggestedRemediation = "Open or update the source workbook, then validate the workbook again.",
                    LinkSource = source,
                    LinkStatus = "indeterminate"
                });
                continue;
            }

            var linkStatus = MapExternalLinkStatus(status);
            findings.Add(new WorkbookIntegrityFinding
            {
                Code = linkStatus.Code,
                Severity = linkStatus.Severity,
                Category = WorkbookIntegrityCategory.ExternalLink,
                Reliability = WorkbookIntegrityReliability.Deterministic,
                Message = linkStatus.Message,
                SuggestedRemediation = linkStatus.SuggestedRemediation,
                LinkSource = source,
                LinkStatus = linkStatus.Status
            });
        }
    }

    private static ExternalLinkValidationStatus MapExternalLinkStatus(int status) =>
        status switch
        {
            0 => new(
                "external-link-ok",
                "ok",
                WorkbookIntegritySeverity.Information,
                "Excel reports that the external workbook link is healthy.",
                "No action is required."),
            1 => new(
                "external-link-missing-file",
                "missing-file",
                WorkbookIntegritySeverity.Error,
                "The source file for an external workbook link is missing.",
                "Restore the source file, update the link source, or explicitly break the link."),
            2 => new(
                "external-link-missing-sheet",
                "missing-sheet",
                WorkbookIntegritySeverity.Error,
                "The source worksheet for an external workbook link is missing.",
                "Restore the source worksheet, update dependent formulas, or explicitly break the link."),
            3 => new(
                "external-link-old",
                "old",
                WorkbookIntegritySeverity.Warning,
                "The external workbook link may be out of date.",
                "Update the link, then validate the workbook again."),
            4 => new(
                "external-link-source-not-calculated",
                "source-not-calculated",
                WorkbookIntegritySeverity.Warning,
                "The external link source has not been calculated.",
                "Calculate and save the source workbook, update the link, then validate again."),
            5 => new(
                "external-link-indeterminate",
                "indeterminate",
                WorkbookIntegritySeverity.Warning,
                "Excel could not determine the external workbook link status.",
                "Open or update the source workbook, then validate the workbook again."),
            6 => new(
                "external-link-not-started",
                "not-started",
                WorkbookIntegritySeverity.Warning,
                "Excel has not started resolving the external workbook link.",
                "Update the link, then validate the workbook again."),
            7 => new(
                "external-link-invalid-name",
                "invalid-name",
                WorkbookIntegritySeverity.Error,
                "The external workbook link contains an invalid defined name.",
                "Correct the source name or dependent formula."),
            8 => new(
                "external-link-source-not-open",
                "source-not-open",
                WorkbookIntegritySeverity.Information,
                "The external workbook link source is not open.",
                "No action is required unless current source values are needed."),
            9 => new(
                "external-link-source-open",
                "source-open",
                WorkbookIntegritySeverity.Information,
                "The external workbook link source is open.",
                "No action is required."),
            10 => new(
                "external-link-copied-values",
                "copied-values",
                WorkbookIntegritySeverity.Information,
                "Excel is using copied values for the external workbook link.",
                "Update the link if current source values are required."),
            _ => new(
                "external-link-unknown-status",
                $"unknown-{status}",
                WorkbookIntegritySeverity.Warning,
                $"Excel returned unknown external link status {status}.",
                "Update the link, then validate the workbook again.")
        };

    private static void ValidateTables(
        Excel.Worksheet worksheet,
        string worksheetName,
        IntegrityFindingCollector findings,
        CancellationToken cancellationToken)
    {
        Excel.ListObjects? tables = null;
        try
        {
            tables = worksheet.ListObjects;
            int tableCount = tables.Count;
            for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.ListObject? table = null;
                try
                {
                    table = tables.Item[tableIndex];
                    string tableName = table.Name;
                    ValidateTableStructure(
                        worksheetName,
                        table,
                        tableName,
                        findings,
                        cancellationToken);
                    ValidateCalculatedColumns(
                        worksheetName,
                        table,
                        tableName,
                        findings,
                        cancellationToken);
                }
                finally
                {
                    ComUtilities.Release(ref table);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref tables);
        }
    }

    private static void ValidateTableStructure(
        string worksheetName,
        Excel.ListObject table,
        string tableName,
        IntegrityFindingCollector findings,
        CancellationToken cancellationToken)
    {
        Excel.Range? tableRange = null;
        Excel.Range? headerRange = null;
        Excel.Range? dataBodyRange = null;
        Excel.ListColumns? columns = null;
        Excel.ListRows? rows = null;
        try
        {
            tableRange = table.Range;
            columns = table.ListColumns;
            rows = table.ListRows;
            var (_, tableColumnCount) = GetRangeDimensions(tableRange);
            int listColumnCount = columns.Count;
            if (tableColumnCount != listColumnCount)
            {
                findings.Add(CreateTableFinding(
                    tableName,
                    worksheetName,
                    "table-column-count-mismatch",
                    WorkbookIntegritySeverity.Error,
                    WorkbookIntegrityCategory.TableStructure,
                    $"Table range has {tableColumnCount} columns but Excel reports {listColumnCount} table columns.",
                    "Resize or recreate the table so its range and column collection agree."));
            }

            if (!table.ShowHeaders)
            {
                findings.Add(CreateTableFinding(
                    tableName,
                    worksheetName,
                    "table-headers-hidden",
                    WorkbookIntegritySeverity.Warning,
                    WorkbookIntegrityCategory.TableHeader,
                    "Table headers are hidden.",
                    "Show the table header row before delivering the workbook."));
            }
            else
            {
                headerRange = table.HeaderRowRange;
                if (headerRange is null)
                {
                    findings.Add(CreateTableFinding(
                        tableName,
                        worksheetName,
                        "table-header-range-missing",
                        WorkbookIntegritySeverity.Error,
                        WorkbookIntegrityCategory.TableStructure,
                        "Excel reports visible table headers but no header range.",
                        "Resize or recreate the table."));
                }
                else
                {
                    var (_, headerColumnCount) = GetRangeDimensions(headerRange);
                    if (headerColumnCount != listColumnCount)
                    {
                        findings.Add(CreateTableFinding(
                            tableName,
                            worksheetName,
                            "table-header-count-mismatch",
                            WorkbookIntegritySeverity.Error,
                            WorkbookIntegrityCategory.TableStructure,
                            $"Table header range has {headerColumnCount} cells but Excel reports {listColumnCount} columns.",
                            "Resize or recreate the table so every column has one header."));
                    }

                    ValidateTableHeaders(
                        worksheetName,
                        tableName,
                        headerRange,
                        headerColumnCount,
                        findings,
                        cancellationToken);
                }
            }

            dataBodyRange = table.DataBodyRange;
            int listRowCount = rows.Count;
            if (dataBodyRange is null)
            {
                if (listRowCount != 0)
                {
                    findings.Add(CreateTableFinding(
                        tableName,
                        worksheetName,
                        "table-data-range-missing",
                        WorkbookIntegritySeverity.Error,
                        WorkbookIntegrityCategory.TableStructure,
                        $"Excel reports {listRowCount} table rows but no data range.",
                        "Resize or recreate the table."));
                }
            }
            else
            {
                var (dataRowCount, dataColumnCount) = GetRangeDimensions(dataBodyRange);
                if (dataRowCount != listRowCount || dataColumnCount != listColumnCount)
                {
                    findings.Add(CreateTableFinding(
                        tableName,
                        worksheetName,
                        "table-data-dimensions-mismatch",
                        WorkbookIntegritySeverity.Error,
                        WorkbookIntegrityCategory.TableStructure,
                        $"Table data range is {dataRowCount} by {dataColumnCount}, but Excel reports {listRowCount} rows and {listColumnCount} columns.",
                        "Resize or recreate the table so its data range matches its row and column collections."));
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref rows);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref dataBodyRange);
            ComUtilities.Release(ref headerRange);
            ComUtilities.Release(ref tableRange);
        }
    }

    private static void ValidateTableHeaders(
        string worksheetName,
        string tableName,
        Excel.Range headerRange,
        int headerColumnCount,
        IntegrityFindingCollector findings,
        CancellationToken cancellationToken)
    {
        object values = headerRange.Value2;
        var seenHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int columnOffset = 0; columnOffset < headerColumnCount; columnOffset++)
        {
            if ((columnOffset & 255) == 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            object? value = GetMatrixValue(values, 0, columnOffset);
            string address = GetCellAddress(headerRange.Row, headerRange.Column + columnOffset);
            if (ExcelErrorMapper.TryGet(value, out int errorCode, out var error))
            {
                var finding = CreateTableFinding(
                    tableName,
                    worksheetName,
                    "table-header-error",
                    WorkbookIntegritySeverity.Error,
                    WorkbookIntegrityCategory.TableHeader,
                    $"Table header contains {error.Name}.",
                    "Replace the header error with a unique text label.");
                finding.CellAddress = address;
                finding.ErrorName = error.Name;
                finding.ErrorCode = errorCode;
                findings.Add(finding);
                continue;
            }

            string header = Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
            if (header.Length == 0)
            {
                var finding = CreateTableFinding(
                    tableName,
                    worksheetName,
                    "table-header-empty",
                    WorkbookIntegritySeverity.Error,
                    WorkbookIntegrityCategory.TableHeader,
                    "Table header is empty.",
                    "Assign a unique text label to every table column.");
                finding.CellAddress = address;
                findings.Add(finding);
            }
            else if (!seenHeaders.Add(header))
            {
                var finding = CreateTableFinding(
                    tableName,
                    worksheetName,
                    "table-header-duplicate",
                    WorkbookIntegritySeverity.Error,
                    WorkbookIntegrityCategory.TableHeader,
                    $"Table header '{header}' is duplicated.",
                    "Assign a unique text label to every table column.");
                finding.CellAddress = address;
                finding.ColumnName = header;
                findings.Add(finding);
            }
        }
    }

    private static void ValidateCalculatedColumns(
        string worksheetName,
        Excel.ListObject table,
        string tableName,
        IntegrityFindingCollector findings,
        CancellationToken cancellationToken)
    {
        Excel.ListColumns? columns = null;
        try
        {
            columns = table.ListColumns;
            int columnCount = columns.Count;
            for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.ListColumn? column = null;
                Excel.Range? dataRange = null;
                try
                {
                    column = columns.Item[columnIndex];
                    string columnName = column.Name;
                    dataRange = column.DataBodyRange;
                    if (dataRange is null)
                    {
                        continue;
                    }

                    var (rowCount, dataColumnCount) = GetRangeDimensions(dataRange);
                    if (rowCount < 2 || dataColumnCount != 1)
                    {
                        continue;
                    }

                    object rawFormulas = dataRange.Formula2R1C1;
                    var formulas = new string?[rowCount];
                    var formulaCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    int formulaCellCount = 0;
                    for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
                    {
                        if ((rowOffset & 1023) == 0)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        string? formula = Convert.ToString(
                            GetMatrixValue(rawFormulas, rowOffset, 0),
                            CultureInfo.InvariantCulture);
                        if (formula?.StartsWith('=') != true)
                        {
                            continue;
                        }

                        formulas[rowOffset] = formula;
                        formulaCellCount++;
                        formulaCounts[formula] = formulaCounts.GetValueOrDefault(formula) + 1;
                    }

                    if (formulaCellCount == 0)
                    {
                        continue;
                    }

                    var dominant = formulaCounts
                        .OrderByDescending(pair => pair.Value)
                        .ThenBy(pair => pair.Key, StringComparer.Ordinal)
                        .First();
                    if (dominant.Value * 2 < rowCount)
                    {
                        continue;
                    }

                    for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
                    {
                        if ((rowOffset & 1023) == 0)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        string? formula = formulas[rowOffset];
                        if (string.Equals(formula, dominant.Key, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        var finding = CreateTableFinding(
                            tableName,
                            worksheetName,
                            "calculated-column-inconsistent",
                            WorkbookIntegritySeverity.Warning,
                            WorkbookIntegrityCategory.CalculatedColumn,
                            $"Table column '{columnName}' differs from its dominant calculated-column formula.",
                            "Confirm the outlier is intentional or restore the column's calculated formula.",
                            WorkbookIntegrityReliability.Heuristic);
                        finding.ColumnName = columnName;
                        finding.CellAddress = GetCellAddress(dataRange.Row + rowOffset, dataRange.Column);
                        finding.Formula = formula;
                        findings.Add(finding);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref dataRange);
                    ComUtilities.Release(ref column);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref columns);
        }
    }

    private static WorkbookIntegrityFinding CreateTableFinding(
        string tableName,
        string worksheetName,
        string code,
        WorkbookIntegritySeverity severity,
        WorkbookIntegrityCategory category,
        string message,
        string remediation,
        WorkbookIntegrityReliability reliability = WorkbookIntegrityReliability.Deterministic) =>
        new()
        {
            Code = code,
            Severity = severity,
            Category = category,
            Reliability = reliability,
            Message = message,
            SuggestedRemediation = remediation,
            SheetName = worksheetName,
            TableName = tableName
        };

    private static void ValidateControlTotals(
        Excel.Workbook workbook,
        List<WorkbookControlTotalExpectation> controlTotals,
        WorkbookIntegrityResult result,
        IntegrityFindingCollector findings,
        CancellationToken cancellationToken)
    {
        foreach (var worksheetExpectations in controlTotals.GroupBy(
                     expectation => expectation.SheetName,
                     StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();
            Excel.Worksheet? worksheet = null;
            try
            {
                worksheet = FindWorksheetIgnoreCase(
                    workbook,
                    worksheetExpectations.Key,
                    cancellationToken)
                    ?? throw new InvalidOperationException(
                        $"Worksheet '{worksheetExpectations.Key}' was not found for control totals.");
                string actualWorksheetName = worksheet.Name;
                if (!result.CheckedWorksheets.Contains(actualWorksheetName, StringComparer.OrdinalIgnoreCase))
                {
                    result.CheckedWorksheets.Add(actualWorksheetName);
                }

                foreach (var expectation in worksheetExpectations)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Excel.Range? range = null;
                    try
                    {
                        try
                        {
                            range = worksheet.Range[expectation.CellAddress];
                        }
                        catch (COMException exception) when (exception.HResult == ExcelNoMatchingCellsHResult)
                        {
                            throw new ArgumentException(
                                $"Control total at '{expectation.SheetName}!{expectation.CellAddress}' does not contain a valid cell address.",
                                nameof(controlTotals),
                                exception);
                        }

                        var (rowCount, columnCount) = GetRangeDimensions(range);
                        if (rowCount != 1 || columnCount != 1)
                        {
                            throw new ArgumentException(
                                $"Control total '{expectation.SheetName}!{expectation.CellAddress}' must identify one cell.",
                                nameof(controlTotals));
                        }

                        object? actualValue = range.Value2;
                        if (ExcelErrorMapper.TryGet(actualValue, out int errorCode, out var error))
                        {
                            findings.Add(CreateControlTotalFinding(
                                expectation,
                                actualWorksheetName,
                                range,
                                "control-total-formula-error",
                                $"{error.Name} prevents the control total from being compared.",
                                error.Name,
                                error.Name,
                                errorCode));
                            continue;
                        }

                        if (!TryConvertFiniteNumber(actualValue, out double actualNumber))
                        {
                            findings.Add(CreateControlTotalFinding(
                                expectation,
                                actualWorksheetName,
                                range,
                                "control-total-non-numeric",
                                "Control-total cell does not contain a finite numeric value.",
                                actualValue,
                                errorName: null));
                            continue;
                        }

                        double expectedValue = expectation.ExpectedValue!.Value;
                        if (Math.Abs(actualNumber - expectedValue) > expectation.Tolerance)
                        {
                            findings.Add(CreateControlTotalFinding(
                                expectation,
                                actualWorksheetName,
                                range,
                                "control-total-mismatch",
                                $"Control total is {actualNumber.ToString("R", CultureInfo.InvariantCulture)}; expected {expectedValue.ToString("R", CultureInfo.InvariantCulture)} within {expectation.Tolerance.ToString("R", CultureInfo.InvariantCulture)}.",
                                actualNumber,
                                errorName: null));
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref range);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref worksheet);
            }
        }
    }

    private static WorkbookIntegrityFinding CreateControlTotalFinding(
        WorkbookControlTotalExpectation expectation,
        string worksheetName,
        Excel.Range range,
        string code,
        string message,
        object? actualValue,
        string? errorName,
        int? errorCode = null) =>
        new()
        {
            Code = code,
            Severity = WorkbookIntegritySeverity.Error,
            Category = WorkbookIntegrityCategory.ControlTotal,
            Reliability = WorkbookIntegrityReliability.Deterministic,
            Message = message,
            SuggestedRemediation = "Review the source data or update the caller-supplied expected control total.",
            SheetName = worksheetName,
            CellAddress = GetCellAddress(range.Row, range.Column),
            ErrorName = errorName,
            ErrorCode = errorCode,
            ExpectedValue = expectation.ExpectedValue,
            ActualValue = actualValue,
            Tolerance = expectation.Tolerance
        };

    private static Excel.Worksheet? FindWorksheetIgnoreCase(
        Excel.Workbook workbook,
        string worksheetName,
        CancellationToken cancellationToken)
    {
        Excel.Sheets? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;
            int worksheetCount = worksheets.Count;
            for (int index = 1; index <= worksheetCount; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.Worksheet? worksheet = null;
                try
                {
                    worksheet = (Excel.Worksheet)worksheets.Item[index];
                    if (string.Equals(worksheet.Name, worksheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        var match = worksheet;
                        worksheet = null;
                        return match;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref worksheet);
                }
            }

            return null;
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }
    }

    private static bool TryConvertFiniteNumber(object? value, out double number)
    {
        if (value is not (byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal))
        {
            number = default;
            return false;
        }

        number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
        return double.IsFinite(number);
    }

    private static (int Rows, int Columns) GetRangeDimensions(Excel.Range range)
    {
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        try
        {
            rows = range.Rows;
            columns = range.Columns;
            return (Convert.ToInt32(rows.Count, CultureInfo.InvariantCulture),
                Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture));
        }
        finally
        {
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
        }
    }

    private static object? GetMatrixValue(object value, int rowOffset, int columnOffset)
    {
        if (value is not object[,] matrix)
        {
            return rowOffset == 0 && columnOffset == 0 ? value : null;
        }

        return matrix[
            matrix.GetLowerBound(0) + rowOffset,
            matrix.GetLowerBound(1) + columnOffset];
    }

    private static string GetCellAddress(int row, int column)
    {
        string columnName = string.Empty;
        while (column > 0)
        {
            column--;
            columnName = Convert.ToChar('A' + (column % 26), CultureInfo.InvariantCulture) + columnName;
            column /= 26;
        }

        return columnName + row.ToString(CultureInfo.InvariantCulture);
    }

    private readonly record struct ExternalLinkValidationStatus(
        string Code,
        string Status,
        WorkbookIntegritySeverity Severity,
        string Message,
        string SuggestedRemediation);

    private sealed class IntegrityFindingCollector
    {
        private readonly Dictionary<(WorkbookIntegritySeverity Severity, WorkbookIntegrityCategory Category), int> _counts = [];
        private readonly List<WorkbookIntegrityFinding> _retainedFindings = [];
        private readonly WorkbookIntegrityResult _result;
        private readonly int _maxFindings;

        internal IntegrityFindingCollector(WorkbookIntegrityResult result, int maxFindings)
        {
            _result = result;
            _maxFindings = maxFindings;
        }

        internal void Add(WorkbookIntegrityFinding finding)
        {
            var key = (finding.Severity, finding.Category);
            _counts[key] = _counts.GetValueOrDefault(key) + 1;
            _result.FindingCount++;
            switch (finding.Severity)
            {
                case WorkbookIntegritySeverity.Error:
                    _result.ErrorCount++;
                    break;
                case WorkbookIntegritySeverity.Warning:
                    _result.WarningCount++;
                    break;
                case WorkbookIntegritySeverity.Information:
                    _result.InformationCount++;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(finding), finding.Severity, "Unknown finding severity.");
            }

            if (_retainedFindings.Count < _maxFindings)
            {
                _retainedFindings.Add(finding);
            }
            else
            {
                _result.FindingsTruncated = true;
            }
        }

        internal void Complete()
        {
            _result.OverallStatus = _result.ErrorCount > 0
                ? WorkbookIntegrityStatus.Failed
                : _result.WarningCount > 0
                    ? WorkbookIntegrityStatus.PassedWithWarnings
                    : WorkbookIntegrityStatus.Passed;

            _result.Groups = _counts
                .OrderBy(pair => pair.Key.Severity)
                .ThenBy(pair => pair.Key.Category)
                .Select(pair => new WorkbookIntegrityFindingGroup
                {
                    Severity = pair.Key.Severity,
                    Category = pair.Key.Category,
                    Count = pair.Value,
                    Findings = _retainedFindings
                        .Where(finding =>
                            finding.Severity == pair.Key.Severity &&
                            finding.Category == pair.Key.Category)
                        .ToList()
                })
                .ToList();
        }
    }
}
