using System;
using System.Linq;

namespace Sbroenne.ExcelMcp.Core.Security;

/// <summary>
/// Provides validation for Excel table names to prevent injection attacks and ensure compliance with Excel naming rules.
/// </summary>
/// <remarks>
/// Excel table name requirements:
/// - Cannot be empty or null
/// - Maximum 255 characters
/// - Cannot contain spaces
/// - Must start with a letter or underscore
/// - Can only contain letters, numbers, underscores, and periods
/// - Cannot use reserved names (Print_Area, Print_Titles, _FilterDatabase, etc.)
/// - Cannot look like cell references (A1, R1C1, etc.)
/// </remarks>
public static class TableNameValidator
{
    /// <summary>
    /// Maximum allowed length for Excel table names
    /// </summary>
    private const int MaxTableNameLength = 255;

    /// <summary>
    /// Reserved names that cannot be used for Excel tables
    /// </summary>
    private static readonly string[] ReservedNames = new[]
    {
        "Print_Area",
        "Print_Titles",
        "_FilterDatabase",
        "Consolidate_Area",
        "Sheet_Title"
    };

    /// <summary>
    /// Validates an Excel table name according to Excel naming rules
    /// </summary>
    /// <param name="tableName">The table name to validate</param>
    /// <param name="parameterName">Name of the parameter for error messages</param>
    /// <returns>The validated table name (trimmed)</returns>
    /// <exception cref="ArgumentException">Thrown if the table name is invalid</exception>
    public static string ValidateTableName(string tableName, string parameterName = "tableName")
    {
        // Null/empty check
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("Table name cannot be null, empty, or whitespace", parameterName);
        }

        // Trim whitespace
        tableName = tableName.Trim();

        // Length check (Excel limit)
        if (tableName.Length > MaxTableNameLength)
        {
            throw new ArgumentException(
                $"Table name too long: {tableName.Length} characters (maximum: {MaxTableNameLength})",
                parameterName);
        }

        // No spaces allowed
        if (tableName.Contains(' '))
        {
            throw new ArgumentException(
                "Table name cannot contain spaces. Use underscores instead.",
                parameterName);
        }

        // Must start with letter or underscore (prevent formula injection)
        if (!char.IsLetter(tableName[0]) && tableName[0] != '_')
        {
            throw new ArgumentException(
                $"Table name must start with a letter or underscore, not '{tableName[0]}'",
                parameterName);
        }

        // Validate characters (letters, numbers, underscores, periods only)
        foreach (char c in tableName)
        {
            if (!char.IsLetterOrDigit(c) && c != '_' && c != '.')
            {
                throw new ArgumentException(
                    $"Table name contains invalid character: '{c}'. Only letters, numbers, underscores, and periods are allowed.",
                    parameterName);
            }
        }

        // Reserved names check (case-insensitive)
        if (ReservedNames.Contains(tableName, StringComparer.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"'{tableName}' is a reserved name and cannot be used for table names",
                parameterName);
        }

        // Check if name looks like a cell reference (e.g., A1, R1C1)
        // This prevents confusion and potential formula injection
        if (LooksLikeCellReference(tableName))
        {
            throw new ArgumentException(
                $"'{tableName}' looks like a cell reference and cannot be used as a table name",
                parameterName);
        }

        return tableName;
    }

    /// <summary>
    /// Validates a table name and returns a tuple indicating success or failure with error message
    /// </summary>
    /// <param name="tableName">The table name to validate</param>
    /// <returns>Tuple with (isValid, errorMessage). errorMessage is null if valid.</returns>
    public static (bool isValid, string? errorMessage) TryValidateTableName(string tableName)
    {
        try
        {
            ValidateTableName(tableName);
            return (true, null);
        }
        catch (ArgumentException ex)
        {
            return (false, ex.Message);
        }
    }

    /// <summary>
    /// Checks if a string looks like an Excel cell reference
    /// </summary>
    /// <param name="name">The name to check</param>
    /// <returns>True if the name looks like a cell reference</returns>
    private static bool LooksLikeCellReference(string name)
    {
        if (string.IsNullOrEmpty(name) || name.Length > 10)
        {
            return false; // Cell references are typically short
        }

        string upper = name.ToUpperInvariant();

        // R1C1-style reference (e.g., R1C1, R10C5)
        if (upper.StartsWith("R") && upper.Contains("C"))
        {
            int cIndex = upper.IndexOf('C');
            
            if (cIndex > 1 && cIndex < upper.Length - 1)
            {
                string rowPart = upper.Substring(1, cIndex - 1);
                string colPart = upper.Substring(cIndex + 1);
                
                if (rowPart.All(char.IsDigit) && rowPart.Length > 0 &&
                    colPart.All(char.IsDigit) && colPart.Length > 0)
                {
                    return true; // R1C1 style
                }
            }
        }

        // A1-style reference (e.g., A1, XFD1048576)
        // Pattern: 1-3 letters (column) followed by 1-7 digits (row), nothing else
        // Excel columns: A-XFD (max 3 letters)
        // Excel rows: 1-1048576 (max 7 digits)
        int letterCount = 0;
        int digitCount = 0;
        bool switchedToDigits = false;

        for (int i = 0; i < name.Length; i++)
        {
            char c = name[i];
            
            if (char.IsLetter(c))
            {
                if (switchedToDigits)
                {
                    return false; // Letters after digits = not a cell reference
                }
                letterCount++;
            }
            else if (char.IsDigit(c))
            {
                if (letterCount == 0)
                {
                    return false; // Starts with digit = not A1 style
                }
                switchedToDigits = true;
                digitCount++;
            }
            else
            {
                return false; // Contains non-alphanumeric = not a cell reference
            }
        }

        // A1-style: must have 1-3 letters and 1-7 digits
        return letterCount >= 1 && letterCount <= 3 && digitCount >= 1 && digitCount <= 7;
    }
}
