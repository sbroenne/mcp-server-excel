using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class RangeFormulaResults
{
    internal static RangeFormulaResult Create(string filePath, string sheetName, string address,
        int startRow, int startColumn, object? formulas, object? values)
    {
        var result = new RangeFormulaResult { FilePath = filePath, SheetName = sheetName, RangeAddress = address };
        if (formulas is object[,] formulaArray && values is object[,] valueArray)
        {
            result.RowCount = formulaArray.GetLength(0);
            result.ColumnCount = formulaArray.GetLength(1);
            if (valueArray.GetLength(0) != result.RowCount || valueArray.GetLength(1) != result.ColumnCount)
                throw new InvalidDataException("Excel returned inconsistent formula and value dimensions.");
            for (var row = 0; row < result.RowCount; row++)
            {
                var formulaRow = new List<string>();
                var valueRow = new List<object?>();
                for (var column = 0; column < result.ColumnCount; column++)
                {
                    var formula = ReturnedFormula(formulaArray[row + formulaArray.GetLowerBound(0), column + formulaArray.GetLowerBound(1)]);
                    formulaRow.Add(formula);
                    valueRow.Add(ConvertErrorForRead(
                        valueArray[row + valueArray.GetLowerBound(0), column + valueArray.GetLowerBound(1)],
                        formula, startRow + row, startColumn + column, result.CellErrors));
                }
                result.Formulas.Add(formulaRow);
                result.Values.Add(valueRow);
            }
        }
        else
        {
            result.RowCount = 1;
            result.ColumnCount = 1;
            var formula = ReturnedFormula(formulas);
            result.Formulas.Add([formula]);
            result.Values.Add([ConvertErrorForRead(values, formula, startRow, startColumn, result.CellErrors)]);
        }
        result.Success = true;
        return result;
    }

    private static string ReturnedFormula(object? value)
    {
        var text = value?.ToString() ?? string.Empty;
        return text.StartsWith('=') ? text : string.Empty;
    }

    internal static object? ConvertErrorForRead(object? value, string formula, int row, int column, List<RangeCellError> errors) =>
        ExcelErrorMapper.TryGet(value, out var code, out var error)
            ? ConvertMappedErrorForRead(value, formula, row, column, errors, code, error) : value;

    internal static string ConvertMappedErrorForRead(object? value, string formula, int row, int column,
        List<RangeCellError> errors, int code, ExcelErrorMapper.ExcelErrorInfo error)
    {
        errors.Add(new RangeCellError
        {
            CellAddress = $"{RangeCommandValidation.ColumnLetter(column)}{row}",
            ErrorName = error.Name,
            Formula = string.IsNullOrEmpty(formula) ? null : formula,
            Row = row,
            Column = column,
            CurrentValue = value,
            ErrorCode = code,
            ErrorMessage = $"{error.Name} - {error.Description}",
            Suggestion = error.Suggestion
        });
        return error.Name;
    }
}
