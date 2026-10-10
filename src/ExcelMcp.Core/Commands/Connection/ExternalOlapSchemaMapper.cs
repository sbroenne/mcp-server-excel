using System.Globalization;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

internal sealed record OlapMemberSearchScope(
    string ConnectionName,
    string HierarchyUniqueName,
    string LevelUniqueName,
    string? SearchText,
    int PageSize);

internal sealed record OlapMemberPage(
    List<ExternalOlapMemberInfo> Members,
    int ScannedCount,
    bool HasMore,
    long? NextAfterOrdinal);

internal static class ExternalOlapSchemaMapper
{
    private const int ContinuationTokenVersion = 1;
    private const int MaximumContinuationTokenLength = 8192;

    private sealed record ContinuationPayload(
        int Version,
        string ConnectionName,
        string HierarchyUniqueName,
        string LevelUniqueName,
        string? SearchText,
        int PageSize,
        long AfterOrdinal);

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
                Ordinal = RequiredInt64(row, "MEMBER_ORDINAL"),
                MemberType = OptionalInt32(row, "MEMBER_TYPE")
            })
            .ToList();
    }

    public static OlapMemberPage SelectMemberPage(
        IReadOnlyList<ExternalOlapMemberInfo> memberRows,
        int pageSize,
        string? searchText,
        int maximumScannedRows,
        long? afterOrdinal)
    {
        var members = new List<ExternalOlapMemberInfo>(pageSize);
        int scannedCount = 0;
        bool hasMore = false;
        long? nextAfterOrdinal = afterOrdinal;

        foreach (var member in memberRows)
        {
            if (searchText is not null && scannedCount == maximumScannedRows)
            {
                hasMore = true;
                break;
            }

            scannedCount++;
            if (searchText is null || MatchesSearch(member, searchText))
            {
                if (members.Count == pageSize)
                {
                    hasMore = true;
                    nextAfterOrdinal = members[^1].Ordinal;
                    break;
                }

                members.Add(member);
            }

            nextAfterOrdinal = member.Ordinal;
        }

        return new OlapMemberPage(members, scannedCount, hasMore, nextAfterOrdinal);
    }

    public static string BuildMembersQuery(
        string hierarchyUniqueName,
        string levelUniqueName,
        long? afterOrdinal,
        int take)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(hierarchyUniqueName);
        ArgumentException.ThrowIfNullOrWhiteSpace(levelUniqueName);
        ArgumentOutOfRangeException.ThrowIfLessThan(take, 1);

        var query = new StringBuilder(
            $"SELECT TOP {take.ToString(CultureInfo.InvariantCulture)} * FROM $SYSTEM.MDSCHEMA_MEMBERS "
            + $"WHERE [HIERARCHY_UNIQUE_NAME] = '{EscapeDmvString(hierarchyUniqueName)}' "
            + $"AND [LEVEL_UNIQUE_NAME] = '{EscapeDmvString(levelUniqueName)}'");
        if (afterOrdinal.HasValue)
        {
            query.Append(" AND [MEMBER_ORDINAL] > ");
            query.Append(afterOrdinal.Value.ToString(CultureInfo.InvariantCulture));
        }

        query.Append(" ORDER BY [MEMBER_ORDINAL] ASC");
        return query.ToString();
    }

    public static string BuildLevelQuery(string levelUniqueName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(levelUniqueName);
        return "SELECT * FROM $SYSTEM.MDSCHEMA_LEVELS "
            + $"WHERE [LEVEL_UNIQUE_NAME] = '{EscapeDmvString(levelUniqueName)}'";
    }

    public static string BuildHierarchyQuery(string hierarchyUniqueName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(hierarchyUniqueName);
        return "SELECT * FROM $SYSTEM.MDSCHEMA_HIERARCHIES "
            + $"WHERE [HIERARCHY_UNIQUE_NAME] = '{EscapeDmvString(hierarchyUniqueName)}'";
    }

    public static string CreateContinuationToken(OlapMemberSearchScope scope, long afterOrdinal)
    {
        var payload = new ContinuationPayload(
            ContinuationTokenVersion,
            scope.ConnectionName,
            scope.HierarchyUniqueName,
            scope.LevelUniqueName,
            scope.SearchText,
            scope.PageSize,
            afterOrdinal);
        var bytes = JsonSerializer.SerializeToUtf8Bytes(payload);
        return Convert.ToBase64String(bytes)
            .TrimEnd('=')
            .Replace('+', '-')
            .Replace('/', '_');
    }

    public static long ReadContinuationToken(string token, OlapMemberSearchScope expectedScope)
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
                || !string.Equals(payload.HierarchyUniqueName, expectedScope.HierarchyUniqueName, StringComparison.Ordinal)
                || !string.Equals(payload.LevelUniqueName, expectedScope.LevelUniqueName, StringComparison.Ordinal)
                || !string.Equals(payload.SearchText, expectedScope.SearchText, StringComparison.Ordinal)
                || payload.PageSize != expectedScope.PageSize)
            {
                throw InvalidContinuationToken();
            }

            return payload.AfterOrdinal;
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

    private static string EscapeDmvString(string value) =>
        value.Replace("'", "''", StringComparison.Ordinal);

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

    private static long RequiredInt64(IReadOnlyDictionary<string, object?> row, string columnName) =>
        OptionalInt64(row, columnName)
        ?? throw new InvalidOperationException(
            $"The OLAP provider's member rowset did not return required column '{columnName}'.");

    private static ArgumentException InvalidContinuationToken() =>
        new(
            "continuationToken is invalid or does not match the requested connection, hierarchy, level, search text, or page size.",
            "continuationToken");

    private static bool MatchesSearch(ExternalOlapMemberInfo member, string searchText) =>
        member.Caption?.Contains(searchText, StringComparison.OrdinalIgnoreCase) == true
        || member.Name?.Contains(searchText, StringComparison.OrdinalIgnoreCase) == true
        || member.UniqueName.Contains(searchText, StringComparison.OrdinalIgnoreCase);
}
