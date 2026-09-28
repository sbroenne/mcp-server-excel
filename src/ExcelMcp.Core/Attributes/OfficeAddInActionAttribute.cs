namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Declares that a source-contract action has an Office.js implementation candidate.
/// The action remains runtime-gated by the exact workbook add-in binding and requirement set.
/// </summary>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class OfficeAddInActionAttribute(string requirementSet, bool mutation) : Attribute
{
    /// <summary>The minimum numbered ExcelApi requirement set.</summary>
    public string RequirementSet { get; } =
        string.IsNullOrWhiteSpace(requirementSet)
            ? throw new ArgumentException("Requirement set is required.", nameof(requirementSet))
            : requirementSet;

    /// <summary>Whether timeout after dispatch can leave workbook state uncertain.</summary>
    public bool Mutation { get; } = mutation;
}
