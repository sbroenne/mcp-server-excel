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
    public void SchemaRequests_UseStandardRowsetsWithRestrictionsInOleDbOlapOrder()
    {
        var dimensions = ExternalOlapSchemaMapper.BuildDimensionRequest("Sales'Cube");
        var hierarchies = ExternalOlapSchemaMapper.BuildHierarchyRequest("Sales'Cube", "[Date].[Calendar]");
        var levels = ExternalOlapSchemaMapper.BuildLevelRequest("Sales'Cube", "[Date].[Calendar].[Month]");
        var members = ExternalOlapSchemaMapper.BuildMembersRequest(
            "Sales'Cube", "[Date].[Calendar]", "[Date].[Calendar].[Month]");

        Assert.Equal(33, dimensions.Schema);
        Assert.Equal([null, null, "Sales'Cube"], dimensions.Restrictions);
        Assert.Equal(34, hierarchies.Schema);
        Assert.Equal([null, null, "Sales'Cube", null, null, "[Date].[Calendar]"], hierarchies.Restrictions);
        Assert.Equal(35, levels.Schema);
        Assert.Equal(
            [null, null, "Sales'Cube", null, null, null, "[Date].[Calendar].[Month]"],
            levels.Restrictions);
        Assert.Equal(38, members.Schema);
        Assert.Equal(
            [null, null, "Sales'Cube", null, "[Date].[Calendar]", "[Date].[Calendar].[Month]"],
            members.Restrictions);
        Assert.Equal(
            [null, null, "Sales'Cube", null, null, null],
            ExternalOlapSchemaMapper.BuildHierarchyRequest("Sales'Cube").Restrictions);
    }

    [Fact]
    public void MapMembers_ProviderWithoutOrdinalLeavesOrdinalUnset()
    {
        var member = Assert.Single(ExternalOlapSchemaMapper.MapMembers(
        [
            Row(("MEMBER_UNIQUE_NAME", "[Date].[Calendar].[Month].&[A]"), ("MEMBER_CAPTION", "January"))
        ]));

        Assert.Null(member.Ordinal);
        Assert.Equal("January", member.Caption);
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
            "SalesCube", "Calendar Cube", "[Date].[Calendar]", "[Date].[Calendar].[Month]", "current", 50);
        var token = ExternalOlapSchemaMapper.CreateContinuationToken(
            scope,
            new OlapMemberPosition(2, "[Date].[Calendar].[Month].&[2024-02]"));

        Assert.DoesNotContain("[Date]", token, StringComparison.Ordinal);
        Assert.Equal(new OlapMemberPosition(2, "[Date].[Calendar].[Month].&[2024-02]"),
            ExternalOlapSchemaMapper.ReadContinuationToken(token, scope));
        Assert.Throws<ArgumentException>(() =>
            ExternalOlapSchemaMapper.ReadContinuationToken(
                token,
                scope with { SearchText = "prior" }));
        Assert.Throws<ArgumentException>(() =>
            ExternalOlapSchemaMapper.ReadContinuationToken(
                token,
                scope with { CubeName = "Another Cube" }));
    }

    [Fact]
    public void SelectMemberPage_ContinuesInProviderOrderWithoutSkippingMembers()
    {
        var members = new[]
        {
            Member("[Date].[Calendar].[Month].&[Z]", "January", 0),
            Member("[Date].[Calendar].[Month].&[B]", "February", 0),
            Member("[Date].[Calendar].[Month].&[A]", "March", 0)
        };

        var first = ExternalOlapSchemaMapper.SelectMemberPage(
            members.Take(3).ToArray(), pageSize: 2, searchText: null, maximumScannedRows: 10);
        var second = ExternalOlapSchemaMapper.SelectMemberPage(
            members.Skip(first.ScannedCount).ToArray(), pageSize: 2, searchText: null, maximumScannedRows: 10);

        Assert.Equal(["January", "February"], first.Members.Select(member => member.Caption));
        Assert.True(first.HasMore);
        Assert.Equal(2, first.ScannedCount);
        Assert.Equal("[Date].[Calendar].[Month].&[B]", first.LastConsumedUniqueName);
        Assert.Equal("March", Assert.Single(second.Members).Caption);
        Assert.False(second.HasMore);
        Assert.All(first.Members.Concat(second.Members), member => Assert.Equal(0, member.Ordinal));
    }

    [Fact]
    public void SelectMemberPage_SearchContinuationResumesAfterBoundedScan()
    {
        var members = new[]
        {
            Member("[Date].[Calendar].[Relative Period].&[A]", "Prior Year", 0),
            Member("[Date].[Calendar].[Relative Period].&[B]", "Prior Quarter", 0),
            Member("[Date].[Calendar].[Relative Period].&[C]", "Current Month", 0),
            Member("[Date].[Calendar].[Relative Period].&[D]", "Next Month", 0)
        };

        var first = ExternalOlapSchemaMapper.SelectMemberPage(
            members.Take(3).ToArray(), pageSize: 1, searchText: "current", maximumScannedRows: 2);
        var second = ExternalOlapSchemaMapper.SelectMemberPage(
            members.Skip(first.ScannedCount).ToArray(), pageSize: 1, searchText: "current", maximumScannedRows: 2);

        Assert.Empty(first.Members);
        Assert.True(first.HasMore);
        Assert.Equal(2, first.ScannedCount);
        Assert.Equal("[Date].[Calendar].[Relative Period].&[B]", first.LastConsumedUniqueName);
        var current = Assert.Single(second.Members);
        Assert.Equal("[Date].[Calendar].[Relative Period].&[C]", current.UniqueName);
        Assert.Equal("Current Month", current.Caption);
        Assert.False(second.HasMore);
    }

    private static Dictionary<string, object?> Row(
        params (string Name, object? Value)[] values) =>
        values.ToDictionary(value => value.Name, value => value.Value, StringComparer.OrdinalIgnoreCase);

    private static ExternalOlapMemberInfo Member(string uniqueName, string caption, long ordinal) =>
        Assert.Single(ExternalOlapSchemaMapper.MapMembers(
        [
            Row(("MEMBER_UNIQUE_NAME", uniqueName),
                ("MEMBER_CAPTION", caption),
                ("MEMBER_NAME", caption),
                ("MEMBER_ORDINAL", ordinal),
                ("MEMBER_TYPE", 1))
        ]));
}
