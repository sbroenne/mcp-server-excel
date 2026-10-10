using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// External OLAP dimensions, hierarchies, and levels exposed by a workbook connection.
/// </summary>
public sealed class ExternalOlapSchemaResult : ResultBase
{
    /// <summary>
    /// Selected workbook connection name.
    /// </summary>
    public string ConnectionName { get; set; } = string.Empty;

    /// <summary>
    /// Dimensions exposed by the connected cube.
    /// </summary>
    public List<ExternalOlapDimensionInfo> Dimensions { get; set; } = [];

    /// <summary>
    /// Hierarchies exposed by the connected cube.
    /// </summary>
    public List<ExternalOlapHierarchyInfo> Hierarchies { get; set; } = [];

    /// <summary>
    /// Levels exposed by the selected hierarchies.
    /// </summary>
    public List<ExternalOlapLevelInfo> Levels { get; set; } = [];
}

/// <summary>
/// External OLAP dimension metadata.
/// </summary>
public sealed class ExternalOlapDimensionInfo
{
    /// <summary>
    /// Provider unique name used to identify the dimension.
    /// </summary>
    public string UniqueName { get; set; } = string.Empty;

    /// <summary>
    /// Display caption, when supplied by the provider.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Caption { get; set; }
}

/// <summary>
/// External OLAP hierarchy metadata.
/// </summary>
public sealed class ExternalOlapHierarchyInfo
{
    /// <summary>
    /// Unique name of the containing dimension.
    /// </summary>
    public string DimensionUniqueName { get; set; } = string.Empty;

    /// <summary>
    /// Provider unique name used to identify the hierarchy.
    /// </summary>
    public string UniqueName { get; set; } = string.Empty;

    /// <summary>
    /// Display caption, when supplied by the provider.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Caption { get; set; }
}

/// <summary>
/// External OLAP level metadata.
/// </summary>
public sealed class ExternalOlapLevelInfo
{
    /// <summary>
    /// Unique name of the containing hierarchy.
    /// </summary>
    public string HierarchyUniqueName { get; set; } = string.Empty;

    /// <summary>
    /// Provider unique name used to identify the level.
    /// </summary>
    public string UniqueName { get; set; } = string.Empty;

    /// <summary>
    /// Display caption, when supplied by the provider.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Caption { get; set; }

    /// <summary>
    /// Ordinal depth of the level, when supplied by the provider.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Number { get; set; }

    /// <summary>
    /// Provider-reported number of members in the level, when available.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? MemberCount { get; set; }
}

/// <summary>
/// A bounded page of members from an external OLAP hierarchy level.
/// </summary>
public sealed class ExternalOlapMemberSearchResult : ResultBase
{
    /// <summary>
    /// Selected workbook connection name.
    /// </summary>
    public string ConnectionName { get; set; } = string.Empty;

    /// <summary>
    /// Unique name of the member's hierarchy.
    /// </summary>
    public string HierarchyUniqueName { get; set; } = string.Empty;

    /// <summary>
    /// Unique name of the member's level.
    /// </summary>
    public string LevelUniqueName { get; set; } = string.Empty;

    /// <summary>
    /// Optional case-insensitive text filter applied to member names and captions.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SearchText { get; set; }

    /// <summary>
    /// Members returned in this page.
    /// </summary>
    public List<ExternalOlapMemberInfo> Members { get; set; } = [];

    /// <summary>
    /// Opaque token for the next page, or null when there are no more members.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ContinuationToken { get; set; }

    /// <summary>
    /// Number of members returned in this page.
    /// </summary>
    public int ReturnedCount { get; set; }

    /// <summary>
    /// Provider-reported total level members when available and no text filter is used.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? TotalCount { get; set; }

    /// <summary>
    /// Number of level members not included in this page, when the provider reports a total.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? OmittedCount { get; set; }

    /// <summary>
    /// Number of provider rows examined to produce this page.
    /// </summary>
    public int ScannedCount { get; set; }
}

/// <summary>
/// External OLAP member metadata.
/// </summary>
public sealed class ExternalOlapMemberInfo
{
    /// <summary>
    /// Provider unique name used to identify the member.
    /// </summary>
    public string UniqueName { get; set; } = string.Empty;

    /// <summary>
    /// Display caption, when supplied by the provider.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Caption { get; set; }

    /// <summary>
    /// Provider member name, when supplied.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; set; }

    /// <summary>
    /// Provider member ordinal, when supplied; it may not be unique and is not used for paging.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? Ordinal { get; set; }

    /// <summary>
    /// Provider member type, when supplied.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? MemberType { get; set; }
}
