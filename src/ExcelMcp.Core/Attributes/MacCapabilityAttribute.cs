namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>macOS implementation tier selected for a public Excel operation.</summary>
public enum MacCapabilityTier
{
    /// <summary>Built into the Apple Events backend without an optional helper.</summary>
    Native,
    /// <summary>Optional Office.js add-in.</summary>
    OfficeAddIn,
    /// <summary>No selected macOS implementation tier.</summary>
    Unsupported,
    /// <summary>Optional native window-capture helper with explicit screen permission.</summary>
    OptionalNativeHelper
}

/// <summary>Current implementation state of a public operation on macOS.</summary>
public enum MacImplementationStatus
{
    /// <summary>The selected contract is implemented and enabled.</summary>
    Implemented,
    /// <summary>A constrained contract variant is implemented or evidence remains incomplete.</summary>
    Partial,
    /// <summary>The selected tier cannot yet meet the contract.</summary>
    Blocked,
    /// <summary>No action-specific implementation has been tested.</summary>
    NotTested
}

/// <summary>
/// Declares the macOS capability decision for a generated command category or action.
/// Method metadata overrides interface metadata.
/// </summary>
[AttributeUsage(
    AttributeTargets.Interface | AttributeTargets.Method | AttributeTargets.Field,
    AllowMultiple = false,
    Inherited = false)]
public sealed class MacCapabilityAttribute(
    MacCapabilityTier tier,
    MacImplementationStatus status,
    bool isAvailable) : Attribute
{
    /// <summary>Gets the selected macOS implementation tier.</summary>
    public MacCapabilityTier Tier { get; } = tier;

    /// <summary>Gets the current implementation status.</summary>
    public MacImplementationStatus Status { get; } = status;

    /// <summary>Gets whether production macOS routing currently permits the action.</summary>
    public bool IsAvailable { get; } = isAvailable;

    /// <summary>Gets or sets the action-specific implementation or test evidence.</summary>
    public string? Evidence { get; set; }

    /// <summary>Gets or sets the Excel or API version associated with the evidence.</summary>
    public string? ExcelApiVersion { get; set; }

    /// <summary>Gets or sets the precise reason the action remains unavailable.</summary>
    public string? Blocker { get; set; }
}
