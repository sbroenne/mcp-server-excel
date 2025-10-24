using System;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Security;

/// <summary>
/// Provides validation for Excel ranges to prevent denial-of-service attacks and invalid operations
/// </summary>
public static class RangeValidator
{
    /// <summary>
    /// Maximum allowed cell count in a range to prevent DoS attacks
    /// Default: 1 million cells (e.g., 1000 rows x 1000 columns)
    /// </summary>
    private const long MaxCellCount = 1_000_000;

    /// <summary>
    /// Validates an Excel range COM object to ensure it's safe to process
    /// </summary>
    /// <param name="rangeObj">The Excel range COM object to validate</param>
    /// <param name="maxCells">Maximum allowed cell count (default: 1,000,000)</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <exception cref="ArgumentNullException">Thrown if rangeObj is null</exception>
    /// <exception cref="ArgumentException">Thrown if range is invalid or too large</exception>
    public static void ValidateRange(dynamic rangeObj, long maxCells = MaxCellCount, string parameterName = "range")
    {
        if (rangeObj == null)
        {
            throw new ArgumentNullException(parameterName, "Range object cannot be null");
        }

        try
        {
            int rowCount = rangeObj.Rows.Count;
            int colCount = rangeObj.Columns.Count;

            // Validate positive dimensions
            if (rowCount < 1 || colCount < 1)
            {
                throw new ArgumentException(
                    $"Range must contain at least one cell. Found {rowCount} rows and {colCount} columns.",
                    parameterName);
            }

            // Calculate total cell count (prevent overflow)
            long cellCount = (long)rowCount * (long)colCount;

            // Prevent DoS with huge ranges
            if (cellCount > maxCells)
            {
                throw new ArgumentException(
                    $"Range too large: {cellCount:N0} cells (maximum: {maxCells:N0}). " +
                    $"Dimensions: {rowCount:N0} rows × {colCount:N0} columns. " +
                    "This limit prevents denial-of-service attacks from processing extremely large ranges.",
                    parameterName);
            }
        }
        catch (COMException ex)
        {
            throw new ArgumentException(
                $"Invalid range object: {ex.Message} (HRESULT: 0x{ex.HResult:X8})",
                parameterName,
                ex);
        }
        catch (ArgumentException)
        {
            // Re-throw ArgumentException as-is
            throw;
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"Error validating range: {ex.Message}",
                parameterName,
                ex);
        }
    }

    /// <summary>
    /// Validates a range and returns a tuple indicating success or failure with error message
    /// </summary>
    /// <param name="rangeObj">The Excel range COM object to validate</param>
    /// <param name="maxCells">Maximum allowed cell count (default: 1,000,000)</param>
    /// <returns>Tuple with (isValid, errorMessage, rowCount, colCount, cellCount). errorMessage is null if valid.</returns>
    public static (bool isValid, string? errorMessage, int rowCount, int colCount, long cellCount) TryValidateRange(
        dynamic rangeObj, 
        long maxCells = MaxCellCount)
    {
        if (rangeObj == null)
        {
            return (false, "Range object is null", 0, 0, 0);
        }

        try
        {
            int rowCount = rangeObj.Rows.Count;
            int colCount = rangeObj.Columns.Count;
            long cellCount = (long)rowCount * (long)colCount;

            if (rowCount < 1 || colCount < 1)
            {
                return (false, $"Range has invalid dimensions: {rowCount} rows × {colCount} columns", rowCount, colCount, cellCount);
            }

            if (cellCount > maxCells)
            {
                return (false, 
                    $"Range too large: {cellCount:N0} cells exceeds maximum of {maxCells:N0}",
                    rowCount, colCount, cellCount);
            }

            return (true, null, rowCount, colCount, cellCount);
        }
        catch (Exception ex)
        {
            return (false, $"Error validating range: {ex.Message}", 0, 0, 0);
        }
    }

    /// <summary>
    /// Validates a range address string format (e.g., "A1:B10")
    /// </summary>
    /// <param name="rangeAddress">The range address to validate</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <returns>The validated range address (trimmed)</returns>
    /// <exception cref="ArgumentException">Thrown if range address format is invalid</exception>
    public static string ValidateRangeAddress(string rangeAddress, string parameterName = "rangeAddress")
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
        {
            throw new ArgumentException("Range address cannot be null, empty, or whitespace", parameterName);
        }

        rangeAddress = rangeAddress.Trim();

        // Basic validation - should contain colon for range (A1:B10) or be single cell (A1)
        // More detailed validation happens when Excel parses it
        if (rangeAddress.Length > 255)
        {
            throw new ArgumentException(
                $"Range address too long: {rangeAddress.Length} characters (maximum: 255)",
                parameterName);
        }

        // Check for obviously invalid characters
        foreach (char c in rangeAddress)
        {
            // Allow letters, digits, colon, dollar sign (for absolute references), exclamation (for sheet names)
            if (!char.IsLetterOrDigit(c) && c != ':' && c != '$' && c != '!' && c != '_' && c != '.')
            {
                throw new ArgumentException(
                    $"Range address contains invalid character: '{c}'",
                    parameterName);
            }
        }

        return rangeAddress;
    }
}
