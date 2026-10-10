using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "ExternalOlapSchema")]
[Trait("RequiresExcel", "false")]
public sealed class ExternalOlapSchemaTests
{
    [Fact]
    public void MapSchema_CalendarHierarchyIncludesRelativePeriodUniqueNamesAndCaptions()
    {
        var dimensions = new[]
        {
            Row(("DIMENSION_UNIQUE_NAME", "[Date]"), ("DIMENSION_CAPTION", "Date"))
        };
        var hierarchies = new[]
        {
            Row(("DIMENSION_UNIQUE_NAME", "[Date]"),
                ("HIERARCHY_UNIQUE_NAME", "[Date].[Calendar]"),
                ("HIERARCHY_CAPTION", "Calendar"))
        };
        var levels = new[]
        {
            Row(("HIERARCHY_UNIQUE_NAME", "[Date].[Calendar]"),
                ("LEVEL_UNIQUE_NAME", "[Date].[Calendar].[Relative Period]"),
                ("LEVEL_CAPTION", "Relative Period"),
                ("LEVEL_NUMBER", 1),
                ("LEVEL_CARDINALITY", 3))
        };
        var memberRows = new[]
        {
            Row(("MEMBER_UNIQUE_NAME", "[Date].[Calendar].[Relative Period].&[Current Month]"),
                ("MEMBER_CAPTION", "Current Month"),
                ("MEMBER_NAME", "Current Month"),
                ("MEMBER_ORDINAL", 0),
                ("MEMBER_TYPE", 1))
        };

        var schema = ExternalOlapSchemaMapper.MapSchema(
            "SalesCube", dimensions, hierarchies, levels);

        var dimension = Assert.Single(schema.Dimensions);
        Assert.Equal("[Date]", dimension.UniqueName);
        Assert.Equal("Date", dimension.Caption);
        var hierarchy = Assert.Single(schema.Hierarchies);
        Assert.Equal("[Date].[Calendar]", hierarchy.UniqueName);
        Assert.Equal("Calendar", hierarchy.Caption);
        var level = Assert.Single(schema.Levels);
        Assert.Equal("[Date].[Calendar].[Relative Period]", level.UniqueName);
        Assert.Equal(3, level.MemberCount);
        var member = Assert.Single(ExternalOlapSchemaMapper.MapMembers(memberRows));
        Assert.Equal("[Date].[Calendar].[Relative Period].&[Current Month]", member.UniqueName);
        Assert.Equal("Current Month", member.Caption);
        Assert.NotEqual(member.UniqueName, member.Caption);
    }

    [Fact]
    public void BuildMembersQuery_UsesEscapedFiltersAndBoundedContinuation()
    {
        var query = ExternalOlapSchemaMapper.BuildMembersQuery(
            "[Date].[Calendar's]",
            "[Date].[Calendar's].[Month]",
            afterOrdinal: 24,
            take: 51);

        Assert.Contains("SELECT TOP 51 *", query, StringComparison.Ordinal);
        Assert.Contains("[HIERARCHY_UNIQUE_NAME] = '[Date].[Calendar''s]'", query, StringComparison.Ordinal);
        Assert.Contains("[LEVEL_UNIQUE_NAME] = '[Date].[Calendar''s].[Month]'", query, StringComparison.Ordinal);
        Assert.Contains("[MEMBER_ORDINAL] > 24", query, StringComparison.Ordinal);
        Assert.Contains("ORDER BY [MEMBER_ORDINAL]", query, StringComparison.Ordinal);
    }

    [Fact]
    public void MapSchema_NegativeProviderCardinalityMeansUnknown()
    {
        var schema = ExternalOlapSchemaMapper.MapSchema(
            "SalesCube",
            Array.Empty<IReadOnlyDictionary<string, object?>>(),
            Array.Empty<IReadOnlyDictionary<string, object?>>(),
            new[] { Row(("HIERARCHY_UNIQUE_NAME", "[Date].[Calendar]"),
                ("LEVEL_UNIQUE_NAME", "[Date].[Calendar].[Month]"),
                ("LEVEL_CARDINALITY", -1)) });

        Assert.Null(Assert.Single(schema.Levels).MemberCount);
    }

    [Fact]
    public void MapSchema_FiltersHierarchyAndItsLevels()
    {
        var dimensions = new[]
        {
            Row(("DIMENSION_UNIQUE_NAME", "[Date]"), ("DIMENSION_CAPTION", "Date")),
            Row(("DIMENSION_UNIQUE_NAME", "[Sales]"), ("DIMENSION_CAPTION", "Sales"))
        };
        var hierarchies = new[]
        {
            Row(("DIMENSION_UNIQUE_NAME", "[Date]"),
                ("HIERARCHY_UNIQUE_NAME", "[Date].[Calendar]"),
                ("HIERARCHY_CAPTION", "Calendar")),
            Row(("DIMENSION_UNIQUE_NAME", "[Sales]"),
                ("HIERARCHY_UNIQUE_NAME", "[Sales].[Territory]"),
                ("HIERARCHY_CAPTION", "Territory"))
        };
        var levels = new[]
        {
            Row(("HIERARCHY_UNIQUE_NAME", "[Date].[Calendar]"),
                ("LEVEL_UNIQUE_NAME", "[Date].[Calendar].[Year]"),
                ("LEVEL_CAPTION", "Year")),
            Row(("HIERARCHY_UNIQUE_NAME", "[Sales].[Territory]"),
                ("LEVEL_UNIQUE_NAME", "[Sales].[Territory].[Country]"),
                ("LEVEL_CAPTION", "Country"))
        };

        var schema = ExternalOlapSchemaMapper.MapSchema(
            "SalesCube",
            dimensions,
            hierarchies,
            levels,
            hierarchyUniqueName: "[Date].[Calendar]");

        Assert.Equal("[Date]", Assert.Single(schema.Dimensions).UniqueName);
        Assert.Equal("[Date].[Calendar]", Assert.Single(schema.Hierarchies).UniqueName);
        Assert.Equal("[Date].[Calendar].[Year]", Assert.Single(schema.Levels).UniqueName);
    }

    [Fact]
    public void ContinuationToken_IsOpaqueAndBoundToItsQuery()
    {
        var scope = new OlapMemberSearchScope(
            "SalesCube", "[Date].[Calendar]", "[Date].[Calendar].[Month]", "current", 50);
        var token = ExternalOlapSchemaMapper.CreateContinuationToken(scope, 24);

        Assert.DoesNotContain("[Date]", token, StringComparison.Ordinal);
        Assert.Equal(24, ExternalOlapSchemaMapper.ReadContinuationToken(token, scope));
        Assert.Throws<ArgumentException>(() =>
            ExternalOlapSchemaMapper.ReadContinuationToken(
                token,
                scope with { SearchText = "prior" }));
    }

    [Fact]
    public void SelectMemberPage_ContinuesWithoutSkippingMembers()
    {
        var members = new[]
        {
            Member("[Date].[Calendar].[Month].&[January]", "January", 0),
            Member("[Date].[Calendar].[Month].&[February]", "February", 1),
            Member("[Date].[Calendar].[Month].&[March]", "March", 2)
        };

        var first = ExternalOlapSchemaMapper.SelectMemberPage(
            members, pageSize: 2, searchText: null, maximumScannedRows: 10, afterOrdinal: null);
        var remaining = members.Where(member => member.Ordinal > first.NextAfterOrdinal.GetValueOrDefault()).ToArray();
        var second = ExternalOlapSchemaMapper.SelectMemberPage(
            remaining, pageSize: 2, searchText: null, maximumScannedRows: 10,
            afterOrdinal: first.NextAfterOrdinal);

        Assert.Equal(["January", "February"], first.Members.Select(member => member.Caption));
        Assert.True(first.HasMore);
        Assert.Equal(1L, first.NextAfterOrdinal);
        Assert.Equal("March", Assert.Single(second.Members).Caption);
        Assert.False(second.HasMore);
    }

    [Fact]
    public void SelectMemberPage_SearchContinuationResumesAfterBoundedScan()
    {
        var members = new[]
        {
            Member("[Date].[Calendar].[Relative Period].&[Prior Year]", "Prior Year", 0),
            Member("[Date].[Calendar].[Relative Period].&[Prior Quarter]", "Prior Quarter", 1),
            Member("[Date].[Calendar].[Relative Period].&[Current Month]", "Current Month", 2),
            Member("[Date].[Calendar].[Relative Period].&[Next Month]", "Next Month", 3)
        };

        var first = ExternalOlapSchemaMapper.SelectMemberPage(
            members, pageSize: 1, searchText: "current", maximumScannedRows: 2, afterOrdinal: null);
        var remaining = members.Where(member => member.Ordinal > first.NextAfterOrdinal.GetValueOrDefault()).ToArray();
        var second = ExternalOlapSchemaMapper.SelectMemberPage(
            remaining, pageSize: 1, searchText: "current", maximumScannedRows: 2,
            afterOrdinal: first.NextAfterOrdinal);

        Assert.Empty(first.Members);
        Assert.True(first.HasMore);
        Assert.Equal(1L, first.NextAfterOrdinal);
        var current = Assert.Single(second.Members);
        Assert.Equal("[Date].[Calendar].[Relative Period].&[Current Month]", current.UniqueName);
        Assert.Equal("Current Month", current.Caption);
        Assert.False(second.HasMore);
    }

    private static Dictionary<string, object?> Row(
        params (string Name, object? Value)[] values) =>
        values.ToDictionary(value => value.Name, value => value.Value, StringComparer.OrdinalIgnoreCase);

    private static ExternalOlapMemberInfo Member(string uniqueName, string caption, long ordinal) =>
        new()
        {
            UniqueName = uniqueName,
            Caption = caption,
            Name = caption,
            Ordinal = ordinal,
            MemberType = 1
        };
}
