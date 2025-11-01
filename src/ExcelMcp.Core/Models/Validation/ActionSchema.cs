namespace Sbroenne.ExcelMcp.Core.Models.Validation;

/// <summary>
/// Defines the schema for an action including its parameters
/// Domain-focused, client-agnostic
/// </summary>
public class ActionSchema
{
    public string Domain { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public ActionParameter[] Parameters { get; init; } = Array.Empty<ActionParameter>();
    public string? Description { get; init; }
}
