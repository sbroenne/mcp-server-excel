using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Data validation operations for Excel ranges (partial class)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult ValidateRange(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string validationType,
        string? validationOperator,
        string? formula1,
        string? formula2,
        bool? showInputMessage,
        string? inputTitle,
        string? inputMessage,
        bool? showErrorAlert,
        string? errorStyle,
        string? errorTitle,
        string? errorMessage,
        bool? ignoreBlank,
        bool? showDropdown)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(validationType);
        var xlType = ParseValidationType(validationType);
        var xlOperator = ParseValidationOperator(validationOperator ?? "between");
        var xlAlertStyle = ParseErrorStyle(errorStyle ?? "stop");
        var effectiveShowErrorAlert = showErrorAlert ?? true;

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Validation? validation = null;

            try
            {
                // Get sheet
                sheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get validation object
                validation = range.Validation;

                // Delete existing validation
                validation.Delete();

                // Add validation
                validation.Add(
                    Type: (Excel.XlDVType)xlType,
                    AlertStyle: xlAlertStyle,
                    Operator: xlOperator,
                    Formula1: formula1 ?? "",
                    Formula2: formula2 ?? "");

                // Configure input message
                validation.ShowInput = showInputMessage ?? false;
                if (showInputMessage is true)
                {
                    validation.InputTitle = inputTitle ?? "";
                    validation.InputMessage = inputMessage ?? "";
                }

                // Configure error alert
                if (effectiveShowErrorAlert)
                {
                    validation.ErrorTitle = errorTitle ?? "";
                    validation.ErrorMessage = errorMessage ?? "";
                }
                validation.ShowError = effectiveShowErrorAlert;

                // Configure additional options
                validation.IgnoreBlank = ignoreBlank ?? true;

                if (validationType.Equals("list", StringComparison.OrdinalIgnoreCase))
                {
                    validation.InCellDropdown = showDropdown ?? true;
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref validation!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private static int ParseValidationType(string type)
    {
        return type.ToLowerInvariant() switch
        {
            "any" => 0, // xlValidateInputOnly
            "whole" => 1, // xlValidateWholeNumber
            "decimal" => 2, // xlValidateDecimal
            "list" => 3, // xlValidateList
            "date" => 4, // xlValidateDate
            "time" => 5, // xlValidateTime
            "textlength" => 6, // xlValidateTextLength
            "custom" => 7, // xlValidateCustom
            _ => throw new ArgumentException($"Invalid validation type: {type}")
        };
    }

    private static int ParseValidationOperator(string op)
    {
        return op.ToLowerInvariant() switch
        {
            "between" => 1, // xlBetween
            "notbetween" => 2, // xlNotBetween
            "equal" => 3, // xlEqual
            "notequal" => 4, // xlNotEqual
            "greaterthan" => 5, // xlGreater
            "lessthan" => 6, // xlLess
            "greaterthanorequal" => 7, // xlGreaterEqual
            "lessthanorequal" => 8, // xlLessEqual
            _ => throw new ArgumentException($"Invalid validation operator: {op}")
        };
    }

    private static int ParseErrorStyle(string style)
    {
        return style.ToLowerInvariant() switch
        {
            "stop" => 1, // xlValidAlertStop
            "warning" => 2, // xlValidAlertWarning
            "information" => 3, // xlValidAlertInformation
            _ => throw new ArgumentException($"Invalid error style: {style}")
        };
    }

    /// <inheritdoc />
    public RangeValidationResult GetValidation(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Validation? validation = null;

            try
            {
                // Get sheet
                sheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Try to get validation
                validation = range.Validation;

                int? nativeType = ReadValidationType(() => validation.Type,
                    () => HasValidationInScope(ctx.App, range));
                if (!nativeType.HasValue)
                {
                    return new RangeValidationResult
                    {
                        Success = true,
                        FilePath = batch.WorkbookPath,
                        SheetName = sheetName,
                        RangeAddress = rangeAddress,
                        HasValidation = false
                    };
                }

                // Read all validation properties into local variables first
                // This ensures we're not affected by any COM state changes during object initialization
                var validationType = GetValidationTypeName(nativeType.Value);
                var validationOperator = GetValidationOperatorName(validation.Operator);
                var formula1 = validation.Formula1?.ToString() ?? string.Empty;
                var formula2 = validation.Formula2?.ToString() ?? string.Empty;
                var ignoreBlank = validation.IgnoreBlank;
                var showInputMessage = validation.ShowInput;
                var inputTitle = validation.InputTitle?.ToString() ?? string.Empty;
                var inputMessage = validation.InputMessage?.ToString() ?? string.Empty;
                var showErrorAlert = validation.ShowError;
                var errorStyle = GetErrorStyleName(validation.AlertStyle);
                var errorTitle = validation.ErrorTitle?.ToString() ?? string.Empty;
                var validationErrorMessage = validation.ErrorMessage?.ToString() ?? string.Empty;

                return new RangeValidationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SheetName = sheetName,
                    RangeAddress = rangeAddress,
                    HasValidation = true,
                    ValidationType = validationType,
                    ValidationOperator = validationOperator,
                    Formula1 = formula1,
                    Formula2 = formula2,
                    IgnoreBlank = ignoreBlank,
                    ShowInputMessage = showInputMessage,
                    InputTitle = inputTitle,
                    InputMessage = inputMessage,
                    ShowErrorAlert = showErrorAlert,
                    ErrorStyle = errorStyle,
                    ErrorTitle = errorTitle,
                    ValidationErrorMessage = validationErrorMessage
                };
            }
            finally
            {
                ComUtilities.Release(ref validation!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Accepts Validation.Type's ambiguous no-rule sentinel only when an independent scope-bound validation-cell query corroborates absence; if a rule is found, the original native failure is rethrown.")]
    private static int? ReadValidationType(Func<int> readType, Func<bool> containsValidation)
    {
        try
        {
            return readType();
        }
        catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
        {
            if (containsValidation())
                throw;
            return null;
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Corroborates a prior Validation.Type no-rule sentinel using only SpecialCells(AllValidation)'s documented no-match result; any found cells are intersected with the exact requested scope, and other native failures propagate.")]
    private static bool HasValidationInScope(Excel.Application app, Excel.Range scope)
    {
        Excel.Range? validatedCells = null;
        Excel.Range? intersection = null;
        try
        {
            try
            {
                validatedCells = scope.SpecialCells(Excel.XlCellType.xlCellTypeAllValidation);
            }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
            {
                return false;
            }
            // Single-cell SpecialCells can search the whole used range instead of the requested cell.
            intersection = app.Intersect(scope, validatedCells);
            return intersection is not null;
        }
        finally
        {
            ComUtilities.Release(ref intersection);
            ComUtilities.Release(ref validatedCells);
        }
    }

    /// <inheritdoc />
    public OperationResult RemoveValidation(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Validation? validation = null;

            try
            {
                // Get sheet
                sheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get validation and delete
                validation = range.Validation;
                validation.Delete();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref validation!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private static string GetValidationTypeName(int type)
    {
        return type switch
        {
            0 => "any",
            1 => "whole",
            2 => "decimal",
            3 => "list",
            4 => "date",
            5 => "time",
            6 => "textlength",
            7 => "custom",
            _ => "unknown"
        };
    }

    private static string GetValidationOperatorName(int op)
    {
        return op switch
        {
            1 => "between",
            2 => "notbetween",
            3 => "equal",
            4 => "notequal",
            5 => "greaterthan",
            6 => "lessthan",
            7 => "greaterthanorequal",
            8 => "lessthanorequal",
            _ => "unknown"
        };
    }

    private static string GetErrorStyleName(int style)
    {
        return style switch
        {
            1 => "stop",
            2 => "warning",
            3 => "information",
            _ => "unknown"
        };
    }
}
