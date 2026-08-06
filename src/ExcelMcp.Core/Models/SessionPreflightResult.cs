namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Real-Excel compatibility and workbook-state evidence for an active session.
/// </summary>
public sealed class SessionPreflightResult
{
    /// <summary>Whether preflight completed successfully.</summary>
    public bool Success { get; init; }

    /// <summary>The active session identifier.</summary>
    public string SessionId { get; init; } = string.Empty;

    /// <summary>The workbook opened by the session.</summary>
    public string FilePath { get; init; } = string.Empty;

    /// <summary>The running Excel environment.</summary>
    public ExcelEnvironmentResult Excel { get; init; } = new();

    /// <summary>Feature availability states.</summary>
    public SessionCapabilitiesResult Capabilities { get; init; } = new();

    /// <summary>Workbook mutability constraints.</summary>
    public WorkbookProtectionResult Workbook { get; init; } = new();

    /// <summary>Human-readable restrictions detected during collection.</summary>
    public List<string> Constraints { get; init; } = [];

    /// <summary>UTC time at which evidence was collected.</summary>
    public DateTime CollectedAtUtc { get; init; }
}

/// <summary>Runtime Excel environment values.</summary>
public sealed class ExcelEnvironmentResult
{
    /// <summary>Excel version.</summary>
    public string Version { get; init; } = string.Empty;

    /// <summary>Excel build number.</summary>
    public int Build { get; init; }

    /// <summary>Bitness of the running Excel process.</summary>
    public string Bitness { get; init; } = string.Empty;

    /// <summary>Excel-reported operating system.</summary>
    public string OperatingSystem { get; init; } = string.Empty;

    /// <summary>Excel user-interface locale.</summary>
    public string UiLocale { get; init; } = string.Empty;
}

/// <summary>Capability states for the active Excel session.</summary>
public sealed class SessionCapabilitiesResult
{
    /// <summary>Formula member compatibility.</summary>
    public Formula2CapabilityResult Formula2 { get; init; } = new();

    /// <summary>Python in Excel availability, never inferred from version alone.</summary>
    public CapabilityResult PythonInExcel { get; init; } = new();

    /// <summary>VBA project trust availability.</summary>
    public CapabilityResult VbaTrust { get; init; } = new();

    /// <summary>Power Pivot availability.</summary>
    public CapabilityResult PowerPivot { get; init; } = new();
}

/// <summary>A capability state with no additional metadata.</summary>
public class CapabilityResult
{
    /// <summary>One of supported, unsupported, unavailable, blocked, or notDetermined.</summary>
    public string Status { get; init; } = "notDetermined";
}

/// <summary>Formula2 compatibility and dynamic-array semantics.</summary>
public sealed class Formula2CapabilityResult : CapabilityResult
{
    /// <summary>Whether modern dynamic-array semantics are available.</summary>
    public bool DynamicArrays { get; init; }
}

/// <summary>Workbook protection and read-only state.</summary>
public sealed class WorkbookProtectionResult
{
    /// <summary>Whether the workbook is read-only.</summary>
    public bool ReadOnly { get; init; }

    /// <summary>Whether workbook structure is protected.</summary>
    public bool StructureProtected { get; init; }

    /// <summary>Whether workbook windows are protected.</summary>
    public bool WindowsProtected { get; init; }

    /// <summary>Whether the source file is detected as IRM/AIP-protected.</summary>
    public bool IrmProtected { get; init; }
}
