using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

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
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? validation = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get validation object
                validation = range.Validation;

                // Delete existing validation
                validation.Delete();

                // Parse validation type
                var xlType = ParseValidationType(validationType);
                var xlOperator = ParseValidationOperator(validationOperator ?? "between");
                var xlAlertStyle = ParseErrorStyle(errorStyle ?? "stop");

                // Add validation
                validation.Add(
                    Type: xlType,
                    AlertStyle: xlAlertStyle,
                    Operator: xlOperator,
                    Formula1: formula1 ?? "",
                    Formula2: formula2 ?? "");

                // Configure input message
                if (showInputMessage == true)
                {
                    validation.ShowInput = true;  // MUST set ShowInput=true BEFORE setting title/message
                    validation.InputTitle = inputTitle ?? "";
                    validation.InputMessage = inputMessage ?? "";
                }

                // Configure error alert
                if (showErrorAlert == true)
                {
                    validation.ErrorTitle = errorTitle ?? "";
                    validation.ErrorMessage = errorMessage ?? "";
                    validation.ShowError = true;
                }

                // Configure additional options
                if (ignoreBlank != null)
                {
                    validation.IgnoreBlank = ignoreBlank.Value;
                }

                if (showDropdown != null && validationType.Equals("list", StringComparison.OrdinalIgnoreCase))
                {
                    validation.InCellDropdown = showDropdown.Value;
                }

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to add validation to range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
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
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? validation = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Try to get validation
                validation = range.Validation;

                // Check if validation exists
                var hasValidation = true;
                try
                {
                    var testType = validation.Type;
                }
                catch
                {
                    hasValidation = false;
                }

                if (!hasValidation)
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
                var validationType = GetValidationTypeName(validation.Type);
                var validationOperator = GetValidationOperatorName(validation.Operator);
                var formula1 = validation.Formula1?.ToString() ?? string.Empty;
                var formula2 = validation.Formula2?.ToString() ?? string.Empty;
                var ignoreBlank = validation.IgnoreBlank ?? true;
                var showInputMessage = validation.ShowInput ?? false;
                var inputTitle = validation.InputTitle?.ToString() ?? string.Empty;
                var inputMessage = validation.InputMessage?.ToString() ?? string.Empty;
                var showErrorAlert = validation.ShowError ?? true;
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
            catch (Exception ex)
            {
                return new RangeValidationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to get validation for range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
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

    /// <inheritdoc />
    public OperationResult RemoveValidation(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? validation = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get validation and delete
                validation = range.Validation;
                validation.Delete();

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to remove validation from range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
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

