using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;


namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range formula validation operations
/// Improvement #1: Formula syntax validation with error detection and suggestions
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public RangeFormulaValidationResult ValidateFormulas(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>>? formulas = null, string? formulasFile = null)
    {
        // Resolve formulas from inline parameter or file
        var resolvedFormulas = ParameterTransforms.ResolveFormulasOrFile(formulas, formulasFile);

        var result = new RangeFormulaValidationResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sheetName,
            RangeAddress = rangeAddress,
            Formulas = resolvedFormulas,
            FormulaCount = resolvedFormulas.Sum(row => row.Count),
            IsValid = true
        };

        return batch.Execute((ctx, ct) =>
        {
            var errors = new List<FormulaValidationError>();
            var warnings = new List<FormulaValidationWarning>();
            int validCount = 0;

            // Parse starting cell from rangeAddress (e.g., "B1" or "B1:B2")
            var (startRow, startCol) = ParseCellAddress(rangeAddress.Split(':')[0]);

            int currentRow = startRow;
            foreach (var formulaRow in resolvedFormulas)
            {
                int currentCol = startCol;
                foreach (var formula in formulaRow)
                {
                    // Generate cell address for error reporting
                    string cellAddress = GetCellAddress(currentRow, currentCol);

                    // Skip empty formulas (cells without formulas)
                    if (string.IsNullOrWhiteSpace(formula))
                    {
                        validCount++;
                        currentCol++;
                        continue;
                    }

                    // Validate formula
                    var validationErrors = ValidateSingleFormula(formula, cellAddress, currentRow, currentCol);
                    if (validationErrors.Count == 0)
                    {
                        validCount++;
                    }
                    else
                    {
                        result.IsValid = false;
                        errors.AddRange(validationErrors);
                    }

                    currentCol++;
                }
                currentRow++;
            }

            result.ValidCount = validCount;
            result.ErrorCount = errors.Count;
            if (errors.Count > 0)
            {
                result.Errors = errors;
            }
            if (warnings.Count > 0)
            {
                result.Warnings = warnings;
            }

            result.Success = true;
            return result;
        });
    }

    /// <summary>
    /// Validates a single formula and returns list of validation errors
    /// </summary>
    private static List<FormulaValidationError> ValidateSingleFormula(string formula, string cellAddress, int row, int col)
    {
        var errors = new List<FormulaValidationError>();

        if (!formula.StartsWith('='))
        {
            errors.Add(new FormulaValidationError
            {
                CellAddress = cellAddress,
                Row = row,
                Column = col,
                Formula = formula,
                Message = "Formula must start with '=' character",
                Category = "syntax-error"
            });
            return errors;
        }

        string formulaContent = formula[1..]; // Remove leading =

        // Check for unclosed parentheses
        int openParen = 0;
        foreach (char c in formulaContent)
        {
            if (c == '(') openParen++;
            else if (c == ')') openParen--;
        }
        if (openParen != 0)
        {
            errors.Add(new FormulaValidationError
            {
                CellAddress = cellAddress,
                Row = row,
                Column = col,
                Formula = formula,
                Message = openParen > 0
                    ? $"Missing {openParen} closing parenthesis(es)"
                    : $"Extra {-openParen} closing parenthesis(es)",
                Category = "syntax-error"
            });
        }

        // Check for common undefined functions without namespace
        var excelAddInFunctions = new[] { "GETVM3", "GETAKS", "GETDISK", "GETESAN", "GETANF", "GETMANAGEDDISK4" };
        foreach (var func in excelAddInFunctions)
        {
            // Match function call without XA2. namespace (case-insensitive)
            // Use simpler pattern: just look for function name followed by (
            if (formulaContent.Contains(func + "(", StringComparison.OrdinalIgnoreCase) ||
                formulaContent.Contains(func + " (", StringComparison.OrdinalIgnoreCase))
            {
                // Make sure it's not already prefixed with XA2.
                if (!formulaContent.Contains("XA2." + func, StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add(new FormulaValidationError
                    {
                        CellAddress = cellAddress,
                        Row = row,
                        Column = col,
                        Formula = formula,
                        Message = $"Function '{func}' requires XA2 add-in namespace",
                        Suggestion = $"=XA2.{func}(...)",
                        Category = "undefined-function"
                    });
                }
            }
        }

        // Check for references to non-existent sheets (basic check for common patterns)
        if (Regex.IsMatch(formulaContent, @"![A-Z]"))
        {
            // Basic detection of sheet references - report as potential issue
            // Full validation would require access to actual sheet names in the workbook
            errors.Add(new FormulaValidationError
            {
                CellAddress = cellAddress,
                Row = row,
                Column = col,
                Formula = formula,
                Message = "Formula contains sheet reference which requires verification",
                Category = "invalid-reference"
            });
        }

        return errors;
    }

    /// <summary>
    /// Converts row,col (1-based) to Excel cell address (e.g., "A1", "B5")
    /// </summary>
    private static string GetCellAddress(int row, int col)
    {
        // Convert column number to letter(s)
        string colLetter = "";
        int temp = col;
        while (temp > 0)
        {
            temp--;
            colLetter = (char)('A' + (temp % 26)) + colLetter;
            temp /= 26;
        }
        return colLetter + row;
    }

    /// <summary>
    /// Parses Excel cell address (e.g., "B1", "AA5") to (row, col) tuple
    /// </summary>
    private static (int row, int col) ParseCellAddress(string cellAddress)
    {
        // Extract column letters and row number
        string colPart = "";
        string rowPart = "";

        foreach (char c in cellAddress)
        {
            if (char.IsLetter(c))
                colPart += c;
            else
                rowPart += c;
        }

        // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27)
        int col = 0;
        foreach (char c in colPart)
        {
            col = col * 26 + (c - 'A' + 1);
        }

        int row = string.IsNullOrEmpty(rowPart) ? 1 : int.Parse(rowPart, System.Globalization.CultureInfo.InvariantCulture);
        return (row, col);
    }
}
