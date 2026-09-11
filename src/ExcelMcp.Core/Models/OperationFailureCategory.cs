namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Known operation conditions that cannot be inferred from a generic exception.
/// </summary>
public enum OperationFailureCategory
{
    /// <summary>An argument or workbook format is unsupported.</summary>
    InvalidInput,
    /// <summary>A requested workbook object does not exist.</summary>
    NotFound,
    /// <summary>A requested workbook object already exists.</summary>
    Conflict,
    /// <summary>Access to an operation is blocked.</summary>
    Permissions,
    /// <summary>The workbook lacks a required feature or data.</summary>
    Prerequisite,
    /// <summary>A required external component is unavailable.</summary>
    DependencyUnavailable
}
