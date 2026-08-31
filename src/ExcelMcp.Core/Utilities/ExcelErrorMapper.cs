namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>
/// Canonical names and guidance for Excel formula error values returned through COM.
/// </summary>
internal static class ExcelErrorMapper
{
    internal const int BusyErrorCode = -2146826237;
    internal const int PythonErrorCode = -2146826233;

    internal static bool TryGet(object? value, out int errorCode, out ExcelErrorInfo error)
    {
        if (TryNormalizeErrorCode(value, out int code) && TryGet(code, out error))
        {
            errorCode = code;
            return true;
        }

        errorCode = default;
        error = default;
        return false;
    }

    internal static bool TryNormalizeErrorCode(object? value, out int errorCode)
    {
        if (value is int code)
        {
            errorCode = code;
            return true;
        }

        if (value is double number
            && number >= int.MinValue
            && number <= int.MaxValue
            && number == Math.Truncate(number))
        {
            errorCode = Convert.ToInt32(number);
            return true;
        }

        errorCode = default;
        return false;
    }

    internal static bool TryGet(int errorCode, out ExcelErrorInfo error)
    {
        error = errorCode switch
        {
            -2146826288 => new("#NULL!", "Invalid intersection of ranges", "Check the intersection operator and referenced ranges.", true),
            -2146826281 => new("#DIV/0!", "Division by zero", "Ensure the formula does not divide by zero.", true),
            -2146826273 => new("#VALUE!", "Wrong type of argument", "Check function names and argument types.", true),
            -2146826265 => new("#REF!", "Invalid cell reference", "Check that referenced cells and ranges still exist.", true),
            -2146826259 => new("#NAME?", "Unrecognized formula name", "Check function names and defined names for spelling errors.", true),
            -2146826252 => new("#NUM!", "Invalid numeric value", "Check numeric inputs and supported value ranges.", true),
            -2146826246 => new("#N/A", "Value not available", "Check that the lookup value and source data are available.", true),
            -2146826245 => new("#GETTING_DATA", "Data is still being retrieved", "Wait for external data retrieval to finish, then read the range again.", false),
            -2146826243 => new("#SPILL!", "Dynamic array result cannot spill", "Clear or move cells that block the dynamic array result.", true),
            -2146826242 => new("#CONNECT!", "Connection is not ready", "Check the data connection and retry after it finishes connecting.", false),
            -2146826241 => new("#BLOCKED!", "Required resource is blocked", "Check Excel privacy, security, and connected-service settings.", false),
            -2146826240 => new("#UNKNOWN!", "Excel cannot identify the data type", "Check the linked data type or service that supplies this value.", true),
            -2146826239 => new("#FIELD!", "Referenced data field is unavailable", "Check the field name and linked data type.", true),
            -2146826238 => new("#CALC!", "Excel cannot complete the calculation", "Check the formula for unsupported or empty array calculations.", true),
            BusyErrorCode => new("#BUSY!", "Calculation or connected data is still in progress", "Wait for calculation to finish, then read the range again.", false),
            PythonErrorCode => new("#PYTHON!", "Python code raised an error (syntax or runtime exception)", "Check the Python formula syntax and runtime inputs.", false),
            _ => default
        };

        return error.Name is not null;
    }

    internal static bool IsExcelFormulaError(int errorCode) =>
        TryGet(errorCode, out var error) && error.IsExcelFormulaError;

    internal static string GetMessage(int errorCode) =>
        TryGet(errorCode, out var error)
            ? $"{error.Name} - {error.Description}"
            : $"#ERROR! - Unknown error code {errorCode}";

    internal readonly record struct ExcelErrorInfo(
        string Name,
        string Description,
        string Suggestion,
        bool IsExcelFormulaError);
}
