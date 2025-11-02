namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Defines a named range parameter with optional initial value for bulk creation
/// </summary>
public class NamedRangeDefinition
{
    /// <summary>
    /// Name of the named range parameter
    /// </summary>
    public required string Name { get; set; }
    
    /// <summary>
    /// Cell reference (e.g., "Sheet1!$A$1" or "=Sheet1!$A$1")
    /// </summary>
    public required string Reference { get; set; }
    
    /// <summary>
    /// Optional initial value to set in the cell
    /// </summary>
    public object? Value { get; set; }
}
