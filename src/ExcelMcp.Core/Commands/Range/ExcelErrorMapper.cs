using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

internal static class ExcelErrorMapper
{
    internal static bool TryGetErrorCode(object? value, out int errorCode)
    {
        errorCode = value switch
        {
            int intValue => intValue,
            ErrorWrapper wrapper => wrapper.ErrorCode,
            _ => 0
        };

        return errorCode < 0;
    }

    internal static string GetMessage(int errorCode) =>
        errorCode switch
        {
            -2146826288 => "#NULL! - Invalid intersection of ranges",
            -2146826281 => "#DIV/0! - Division by zero",
            -2147483648 => "#DIV/0! - Division by zero",
            -2146826273 => "#VALUE! - Wrong type of argument",
            -2146826265 => "#REF! - Invalid cell reference",
            -2146826259 => "#NAME? - Unknown function or name",
            -2146826252 => "#NUM! - Invalid numeric value",
            -2146826246 => "#N/A - Value not available",
            -2146826245 => "#GETTING_DATA - Data is still loading",
            -2146826243 => "#SPILL! - Dynamic array result cannot spill",
            -2146826242 => "#CONNECT! - External data connection is not ready",
            -2146826241 => "#BLOCKED! - Required resource is blocked",
            -2146826240 => "#UNKNOWN! - Excel cannot determine the data type",
            -2146826239 => "#FIELD! - Referenced data field is unavailable",
            -2146826238 => "#CALC! - Calculation failed",
            -2146826237 => "#BUSY! - Calculation is still running",
            -2146826236 => "#DATA! - Linked data returned an error",
            -2142019887 => "#N/A - Value not available",
            -2146826233 => "#PYTHON! - Python code raised an error (syntax or runtime exception)",
            _ => $"#ERROR! - Unknown error code {errorCode}"
        };

    internal static string GetSuggestion(int errorCode) =>
        errorCode switch
        {
            -2146826288 => "Check the intersection operator and referenced ranges.",
            -2146826281 => "Ensure the formula does not divide by zero.",
            -2147483648 => "Ensure the formula does not divide by zero.",
            -2146826273 => "Check function names and argument types.",
            -2146826265 => "Check that referenced cells and ranges still exist.",
            -2146826259 => "Check function and named-range spelling.",
            -2146826252 => "Check numeric inputs and supported value ranges.",
            -2146826246 => "Check that the lookup value and source data are available.",
            -2146826245 => "Wait for the external data operation to finish.",
            -2146826243 => "Clear cells that block the dynamic array result.",
            -2146826242 => "Check the external data connection and try again.",
            -2146826241 => "Allow the required resource or review workbook security settings.",
            -2146826240 => "Check the linked data type and its source.",
            -2146826239 => "Check that the referenced data field still exists.",
            -2146826238 => "Check the formula inputs and array dimensions.",
            -2146826237 => "Wait for calculation to finish and read the cell again.",
            -2146826236 => "Refresh or repair the linked data source.",
            -2142019887 => "Check that the lookup value and source data are available.",
            -2146826233 => "Check the Python formula syntax and runtime inputs.",
            _ => "Check the formula syntax and referenced values."
        };
}
