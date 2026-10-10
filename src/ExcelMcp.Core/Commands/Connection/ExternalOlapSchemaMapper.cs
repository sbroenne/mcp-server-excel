using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

internal sealed record OlapMemberSearchScope(
    string ConnectionName,
    string CubeName,
    string HierarchyUniqueName,
    string LevelUniqueName,
    string? SearchText,
    int PageSize);

internal sealed record OlapMemberPage(
    List<ExternalOlapMemberInfo> Members,
    int ScannedCount,
    bool HasMore,
    string? LastConsumedUniqueName);

/// <summary>
/// Position in the provider's member order: the number of rows already consumed and the
/// unique name of the last consumed row, used to detect a changed member list.
/// </summary>
internal sealed record OlapMemberPosition(int Offset, string LastUniqueName);

/// <summary>
/// A standard OLE DB for OLAP schema rowset request for ADO Connection.OpenSchema.
/// </summary>
internal sealed record OlapSchemaRequest(int Schema, object?[] Restrictions);

internal static class ExternalOlapSchemaMapper
{
    private const int ContinuationTokenVersion = 3;
    private const int MaximumContinuationTokenLength = 8192;

    private sealed record ContinuationPayload(
        int Version,
        string ConnectionName,
        string CubeName,
        string HierarchyUniqueName,
        string LevelUniqueName,
        string? SearchText,
        int PageSize,
        int Offset,
        string LastUniqueName);

    public static ExternalOlapSchemaResult MapSchema(
        string connectionName,
        IEnumerable<IReadOnlyDictionary<string, object?>> dimensionRows,
        IEnumerable<IReadOnlyDictionary<string, object?>> hierarchyRows,
        IEnumerable<IReadOnlyDictionary<string, object?>> levelRows,
        string? hierarchyUniqueName = null,
        string? levelUniqueName = null)
    {
        var dimensions = dimensionRows
            .Select(row => new ExternalOlapDimensionInfo
            {
                UniqueName = RequiredText(row, "DIMENSION_UNIQUE_NAME"),
                Caption = OptionalText(row, "DIMENSION_CAPTION")
            })
            .ToList();
        var hierarchies = hierarchyRows
            .Select(row => new ExternalOlapHierarchyInfo
            {
                DimensionUniqueName = RequiredText(row, "DIMENSION_UNIQUE_NAME"),
                UniqueName = RequiredText(row, "HIERARCHY_UNIQUE_NAME"),
                Caption = OptionalText(row, "HIERARCHY_CAPTION")
            })
            .ToList();
        var levels = levelRows
            .Select(row => new ExternalOlapLevelInfo
            {
                HierarchyUniqueName = RequiredText(row, "HIERARCHY_UNIQUE_NAME"),
                UniqueName = RequiredText(row, "LEVEL_UNIQUE_NAME"),
                Caption = OptionalText(row, "LEVEL_CAPTION"),
                Number = OptionalInt32(row, "LEVEL_NUMBER"),
                MemberCount = OptionalNonnegativeInt64(row, "LEVEL_CARDINALITY")
            })
            .ToList();

        if (hierarchyUniqueName is not null
            && !hierarchies.Any(item => string.Equals(item.UniqueName, hierarchyUniqueName, StringComparison.Ordinal)))
        {
            throw new InvalidOperationException($"Hierarchy '{hierarchyUniqueName}' was not found on the selected OLAP connection.");
        }

        if (levelUniqueName is not null)
        {
            var selectedLevels = levels
                .Where(item => string.Equals(item.UniqueName, levelUniqueName, StringComparison.Ordinal))
                .ToList();
            if (selectedLevels.Count == 0)
            {
                throw new InvalidOperationException($"Level '{levelUniqueName}' was not found on the selected OLAP connection.");
            }

            if (hierarchyUniqueName is not null
                && selectedLevels.Any(item =>
                    !string.Equals(item.HierarchyUniqueName, hierarchyUniqueName, StringComparison.Ordinal)))
            {
                throw new InvalidOperationException(
                    $"Level '{levelUniqueName}' does not belong to hierarchy '{hierarchyUniqueName}'.");
            }

            levels = selectedLevels;
            var selectedHierarchyNames = levels
                .Select(item => item.HierarchyUniqueName)
                .ToHashSet(StringComparer.Ordinal);
            hierarchies = hierarchies
                .Where(item => selectedHierarchyNames.Contains(item.UniqueName))
                .ToList();
        }
        else if (hierarchyUniqueName is not null)
        {
            hierarchies = hierarchies
                .Where(item => string.Equals(item.UniqueName, hierarchyUniqueName, StringComparison.Ordinal))
                .ToList();
            levels = levels
                .Where(item => string.Equals(item.HierarchyUniqueName, hierarchyUniqueName, StringComparison.Ordinal))
                .ToList();
        }

        var selectedDimensionNames = hierarchies
            .Select(item => item.DimensionUniqueName)
            .ToHashSet(StringComparer.Ordinal);

        return new ExternalOlapSchemaResult
        {
            ConnectionName = connectionName,
            Dimensions = dimensions
                .Where(item => selectedDimensionNames.Contains(item.UniqueName))
                .ToList(),
            Hierarchies = hierarchies,
            Levels = levels
        };
    }

    public static List<ExternalOlapMemberInfo> MapMembers(
        IEnumerable<IReadOnlyDictionary<string, object?>> memberRows)
    {
        return memberRows
            .Select(row => new ExternalOlapMemberInfo
            {
                UniqueName = RequiredText(row, "MEMBER_UNIQUE_NAME"),
                Caption = OptionalText(row, "MEMBER_CAPTION"),
                Name = OptionalText(row, "MEMBER_NAME"),
                Ordinal = OptionalInt64(row, "MEMBER_ORDINAL"),
                MemberType = OptionalInt32(row, "MEMBER_TYPE")
            })
            .ToList();
    }

    public static OlapMemberPage SelectMemberPage(
        IReadOnlyList<ExternalOlapMemberInfo> memberRows,
        int pageSize,
        string? searchText,
        int maximumScannedRows)
    {
        var members = new List<ExternalOlapMemberInfo>(pageSize);
        int consumedCount = 0;
        bool hasMore = false;
        string? lastConsumedUniqueName = null;

        foreach (var member in memberRows)
        {
            if (searchText is not null && consumedCount == maximumScannedRows)
            {
                hasMore = true;
                break;
            }

            if (searchText is null || MatchesSearch(member, searchText))
            {
                if (members.Count == pageSize)
                {
                    hasMore = true;
                    break;
                }

                members.Add(member);
            }

            consumedCount++;
            lastConsumedUniqueName = member.UniqueName;
        }

        return new OlapMemberPage(members, consumedCount, hasMore, lastConsumedUniqueName);
    }

    // ADO SchemaEnum values for the standard OLE DB for OLAP schema rowsets.
    internal const int AdSchemaDimensions = 33;
    internal const int AdSchemaHierarchies = 34;
    internal const int AdSchemaLevels = 35;
    internal const int AdSchemaMembers = 38;

    // Restrictions follow the OLE DB for OLAP column order; null means "no restriction".
    public static OlapSchemaRequest BuildDimensionRequest(string cubeName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(cubeName);
        return new OlapSchemaRequest(AdSchemaDimensions, [null, null, cubeName]);
    }

    public static OlapSchemaRequest BuildHierarchyRequest(string cubeName, string? hierarchyUniqueName = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(cubeName);
        ThrowIfBlank(hierarchyUniqueName);
        return new OlapSchemaRequest(
            AdSchemaHierarchies,
            [null, null, cubeName, null, null, hierarchyUniqueName]);
    }

    public static OlapSchemaRequest BuildLevelRequest(string cubeName, string? levelUniqueName = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(cubeName);
        ThrowIfBlank(levelUniqueName);
        return new OlapSchemaRequest(
            AdSchemaLevels,
            [null, null, cubeName, null, null, null, levelUniqueName]);
    }

    public static OlapSchemaRequest BuildMembersRequest(
        string cubeName,
        string hierarchyUniqueName,
        string levelUniqueName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(cubeName);
        ArgumentException.ThrowIfNullOrWhiteSpace(hierarchyUniqueName);
        ArgumentException.ThrowIfNullOrWhiteSpace(levelUniqueName);
        return new OlapSchemaRequest(
            AdSchemaMembers,
            [null, null, cubeName, null, hierarchyUniqueName, levelUniqueName]);
    }

    public static string CreateContinuationToken(OlapMemberSearchScope scope, OlapMemberPosition position)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(position.Offset, 1);
        ArgumentException.ThrowIfNullOrWhiteSpace(position.LastUniqueName);
        var payload = new ContinuationPayload(
            ContinuationTokenVersion,
            scope.ConnectionName,
            scope.CubeName,
            scope.HierarchyUniqueName,
            scope.LevelUniqueName,
            scope.SearchText,
            scope.PageSize,
            position.Offset,
            position.LastUniqueName);
        var bytes = JsonSerializer.SerializeToUtf8Bytes(payload);
        return Convert.ToBase64String(bytes)
            .TrimEnd('=')
            .Replace('+', '-')
            .Replace('/', '_');
    }

    public static OlapMemberPosition ReadContinuationToken(string token, OlapMemberSearchScope expectedScope)
    {
        if (string.IsNullOrWhiteSpace(token) || token.Length > MaximumContinuationTokenLength)
        {
            throw InvalidContinuationToken();
        }

        try
        {
            string padded = token.Replace('-', '+').Replace('_', '/');
            padded += new string('=', (4 - padded.Length % 4) % 4);
            byte[] bytes = Convert.FromBase64String(padded);
            var payload = JsonSerializer.Deserialize<ContinuationPayload>(bytes);
            if (payload is null
                || payload.Version != ContinuationTokenVersion
                || !string.Equals(payload.ConnectionName, expectedScope.ConnectionName, StringComparison.Ordinal)
                || !string.Equals(payload.CubeName, expectedScope.CubeName, StringComparison.Ordinal)
                || !string.Equals(payload.HierarchyUniqueName, expectedScope.HierarchyUniqueName, StringComparison.Ordinal)
                || !string.Equals(payload.LevelUniqueName, expectedScope.LevelUniqueName, StringComparison.Ordinal)
                || !string.Equals(payload.SearchText, expectedScope.SearchText, StringComparison.Ordinal)
                || payload.PageSize != expectedScope.PageSize
                || payload.Offset < 1
                || string.IsNullOrWhiteSpace(payload.LastUniqueName))
            {
                throw InvalidContinuationToken();
            }

            return new OlapMemberPosition(payload.Offset, payload.LastUniqueName);
        }
        catch (FormatException)
        {
            throw InvalidContinuationToken();
        }
        catch (JsonException)
        {
            throw InvalidContinuationToken();
        }
    }

    private static void ThrowIfBlank(string? value)
    {
        if (value is not null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(value);
        }
    }

    private static string RequiredText(IReadOnlyDictionary<string, object?> row, string columnName) =>
        OptionalText(row, columnName)
        ?? throw new InvalidOperationException(
            $"The OLAP provider's schema rowset did not return required column '{columnName}'.");

    private static string? OptionalText(IReadOnlyDictionary<string, object?> row, string columnName)
    {
        if (!row.TryGetValue(columnName, out var value) || value is null or DBNull)
        {
            return null;
        }

        return Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    private static int? OptionalInt32(IReadOnlyDictionary<string, object?> row, string columnName)
    {
        if (!row.TryGetValue(columnName, out var value) || value is null or DBNull)
        {
            return null;
        }

        return Convert.ToInt32(value, CultureInfo.InvariantCulture);
    }

    private static long? OptionalInt64(IReadOnlyDictionary<string, object?> row, string columnName)
    {
        if (!row.TryGetValue(columnName, out var value) || value is null or DBNull)
        {
            return null;
        }

        return Convert.ToInt64(value, CultureInfo.InvariantCulture);
    }

    private static long? OptionalNonnegativeInt64(
        IReadOnlyDictionary<string, object?> row,
        string columnName)
    {
        long? value = OptionalInt64(row, columnName);
        return value is >= 0 ? value : null;
    }

    private static ArgumentException InvalidContinuationToken() =>
        new(
            "continuationToken is invalid or does not match the requested connection, cube, hierarchy, level, search text, or page size.",
            "continuationToken");

    private static bool MatchesSearch(ExternalOlapMemberInfo member, string searchText) =>
        member.Caption?.Contains(searchText, StringComparison.OrdinalIgnoreCase) == true
        || member.Name?.Contains(searchText, StringComparison.OrdinalIgnoreCase) == true
        || member.UniqueName.Contains(searchText, StringComparison.OrdinalIgnoreCase);
}
