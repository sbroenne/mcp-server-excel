namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Changes retained after a Power Query create or update fails.
/// Null inspection fields mean the state could not be established.
/// Load configuration does not establish successful data loading.
/// </summary>
public sealed class PowerQueryPartialState
{
    /// <summary>The affected query.</summary>
    public required string QueryName { get; init; }

    /// <summary>Whether create added the query before loading failed.</summary>
    public bool? QueryCreated { get; init; }

    /// <summary>Whether update saved the replacement M code before refresh failed.</summary>
    public bool? MCodeUpdated { get; init; }

    /// <summary>Whether inspection found the query after failure.</summary>
    public bool? QueryExists { get; set; }

    /// <summary>The inspected destination configuration, not the requested mode.</summary>
    public PowerQueryLoadMode? LoadMode { get; set; }

    /// <summary>The inspected worksheet destination, when present.</summary>
    public string? TargetSheet { get; set; }

    /// <summary>Whether inspection found a matching Data Model table.</summary>
    public bool? IsLoadedToDataModel { get; set; }

    /// <summary>An inspection failure, distinct from the primary operation failure.</summary>
    public string? StateInspectionError { get; set; }
}
