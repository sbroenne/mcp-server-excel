using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(251)]
    public void SearchOlapMembers_InvalidPageSizePreservesWorkbook(int pageSize)
    {
        var retained = SeedRetainedConnection();
        var error = Assert.Throws<ArgumentOutOfRangeException>(() =>
            _connections.SearchOlapMembers(_fixture.BatchToken, retained.Name,
                "[Hierarchy]", "[Level]", pageSize: pageSize));

        Assert.Contains("pageSize must be between 1 and 250", error.Message, StringComparison.Ordinal);
        AssertRetainedConnection(retained);
    }

    [ConfiguredExternalOlapFact]
    [Trait("RunType", "OnDemand")]
    public void ExternalOlapSchema_UsesSelectedCubeAndContinuesThroughService()
    {
        string connectionString = Environment.GetEnvironmentVariable(
            "EXCELMCP_TEST_OLAP_CONNECTION_STRING")!;
        string cubeName = Environment.GetEnvironmentVariable("EXCELMCP_TEST_OLAP_CUBE")!;
        string hierarchyUniqueName = Environment.GetEnvironmentVariable(
            "EXCELMCP_TEST_OLAP_HIERARCHY")!;
        string levelUniqueName = Environment.GetEnvironmentVariable(
            "EXCELMCP_TEST_OLAP_LEVEL")!;
        var connectionName = UniqueConnectionName("ExternalOlap");
        var pivotSheetName = _fixture.CreateTestSheet(_fixture.BatchToken);

        try
        {
            CreateOlapPivotTable(connectionString, cubeName, connectionName, pivotSheetName);
            var retained = SeedRetainedConnection();
            var pivots = _fixture.CreateCommands<IPersistentPivotTableCommands>();
            string before = System.Text.Json.JsonSerializer.Serialize(
                RequireSuccess(pivots.ListFields(_fixture.BatchToken, connectionName + "_Pivot")).Fields);
            AssertSchemaAndMemberPaging(connectionName, hierarchyUniqueName, levelUniqueName);
            if (string.Equals(Environment.GetEnvironmentVariable(
                "EXCELMCP_TEST_OLAP_SYNTHETIC_FIXTURE"), "true", StringComparison.OrdinalIgnoreCase))
            {
                AssertSyntheticCalendarMonths(connectionName);
                AssertSyntheticRelativePeriods(connectionName);
            }
            AssertRetainedConnection(retained);
            Assert.Equal(before, System.Text.Json.JsonSerializer.Serialize(
                RequireSuccess(pivots.ListFields(_fixture.BatchToken, connectionName + "_Pivot")).Fields));
        }
        finally
        {
            _fixture.Send("sheet.delete", new { sheetName = pivotSheetName });
            _fixture.ForgetSheet(pivotSheetName);
        }
    }

    private void AssertSchemaAndMemberPaging(
        string connectionName,
        string hierarchyUniqueName,
        string levelUniqueName)
    {
        var allSchema = RequireSuccess(_connections.DiscoverOlapSchema(
            _fixture.BatchToken, connectionName));
        Assert.Equal(connectionName, allSchema.ConnectionName);
        Assert.True(allSchema.Hierarchies.Count > 1, "The test cube must expose unused hierarchies.");
        Assert.All(allSchema.Hierarchies, hierarchy =>
            Assert.Contains(allSchema.Dimensions, dimension => dimension.UniqueName == hierarchy.DimensionUniqueName));
        var schema = RequireSuccess(_connections.DiscoverOlapSchema(
            _fixture.BatchToken,
            connectionName,
            hierarchyUniqueName));
        Assert.Equal(connectionName, schema.ConnectionName);
        Assert.Contains(schema.Hierarchies, hierarchy => hierarchy.UniqueName == hierarchyUniqueName);
        Assert.Contains(schema.Levels, level => level.UniqueName == levelUniqueName);
        Assert.Equal(hierarchyUniqueName, Assert.Single(schema.Hierarchies).UniqueName);
        Assert.All(schema.Levels, level => Assert.Equal(hierarchyUniqueName, level.HierarchyUniqueName));
        var filteredLevel = RequireSuccess(_connections.DiscoverOlapSchema(
            _fixture.BatchToken, connectionName, hierarchyUniqueName, levelUniqueName));
        Assert.Equal(connectionName, filteredLevel.ConnectionName);
        Assert.Equal(levelUniqueName, Assert.Single(filteredLevel.Levels).UniqueName);

        var firstPage = RequireSuccess(_connections.SearchOlapMembers(
            _fixture.BatchToken,
            connectionName,
            hierarchyUniqueName,
            levelUniqueName,
            pageSize: 1));
        var firstMember = Assert.Single(firstPage.Members);
        Assert.Equal(connectionName, firstPage.ConnectionName);
        Assert.Equal(hierarchyUniqueName, firstPage.HierarchyUniqueName);
        Assert.Equal(levelUniqueName, firstPage.LevelUniqueName);
        Assert.Equal(1, firstPage.ReturnedCount);
        Assert.False(string.IsNullOrWhiteSpace(firstPage.ContinuationToken));

        var secondPage = RequireSuccess(_connections.SearchOlapMembers(
            _fixture.BatchToken,
            connectionName,
            hierarchyUniqueName,
            levelUniqueName,
            continuationToken: firstPage.ContinuationToken,
            pageSize: 1));
        var secondMember = Assert.Single(secondPage.Members);
        Assert.Equal(connectionName, secondPage.ConnectionName);
        Assert.Equal(hierarchyUniqueName, secondPage.HierarchyUniqueName);
        Assert.Equal(levelUniqueName, secondPage.LevelUniqueName);
        Assert.Equal(1, secondPage.ReturnedCount);
        Assert.NotEqual(firstMember.UniqueName, secondMember.UniqueName);
        Assert.True(string.CompareOrdinal(firstMember.UniqueName, secondMember.UniqueName) < 0);
        Assert.InRange(firstPage.ScannedCount, 1, 10000);
        Assert.InRange(secondPage.ScannedCount, 1, 10000);
        var wrongQuery = Assert.Throws<ArgumentException>(() =>
            _connections.SearchOlapMembers(_fixture.BatchToken, connectionName, hierarchyUniqueName,
                levelUniqueName, searchText: "changed query", continuationToken: firstPage.ContinuationToken, pageSize: 1));
        Assert.Contains("continuationToken", wrongQuery.Message, StringComparison.Ordinal);
    }

    private void AssertSyntheticRelativePeriods(string connectionName)
    {
        const string hierarchy = "[Sales].[RelativePeriod]";
        const string level = "[Sales].[RelativePeriod].[RelativePeriod]";
        var schema = RequireSuccess(_connections.DiscoverOlapSchema(
            _fixture.BatchToken, connectionName, hierarchy, level));
        Assert.Equal(level, Assert.Single(schema.Levels).UniqueName);
        var members = RequireSuccess(_connections.SearchOlapMembers(
            _fixture.BatchToken, connectionName, hierarchy, level));
        Assert.Equal(["Current Month", "Earlier Months", "Next Month", "Prior Month"],
            members.Members.Select(member => member.Caption));
        Assert.Equal(members.Members.Count, members.ReturnedCount);
        Assert.Null(members.ContinuationToken);
        Assert.All(members.Members, member =>
        {
            Assert.Equal($"{hierarchy}.[ALL].[AllMember].[{member.Caption}]", member.UniqueName);
            Assert.NotEqual(member.Caption, member.UniqueName);
        });
        string? token = null;
        var pagedNames = new List<string>();
        for (int index = 0; index < 4; index++)
        {
            var page = RequireSuccess(_connections.SearchOlapMembers(
                _fixture.BatchToken, connectionName, hierarchy, level, continuationToken: token, pageSize: 1));
            pagedNames.Add(Assert.Single(page.Members).UniqueName);
            token = page.ContinuationToken;
            if (index < 3)
            {
                Assert.False(string.IsNullOrWhiteSpace(token));
            }
        }
        Assert.Null(token);
        Assert.Equal(members.Members.Select(member => member.UniqueName), pagedNames);

        var current = RequireSuccess(_connections.SearchOlapMembers(
            _fixture.BatchToken, connectionName, hierarchy, level, searchText: "current"));
        Assert.Equal("current", current.SearchText);
        Assert.Equal($"{hierarchy}.[ALL].[AllMember].[Current Month]", Assert.Single(current.Members).UniqueName);
        Assert.Null(current.TotalCount);
    }

    private void AssertSyntheticCalendarMonths(string connectionName)
    {
        const string hierarchy = "[Sales].[Calendar]";
        const string level = "[Sales].[Calendar].[Month]";
        const int pageSize = 5;
        var schema = RequireSuccess(_connections.DiscoverOlapSchema(
            _fixture.BatchToken, connectionName, hierarchy, level));
        var dimension = Assert.Single(schema.Dimensions);
        Assert.Equal("[Sales]", dimension.UniqueName);
        Assert.Equal("Sales", dimension.Caption);
        var selectedHierarchy = Assert.Single(schema.Hierarchies);
        Assert.Equal(hierarchy, selectedHierarchy.UniqueName);
        Assert.Equal(dimension.UniqueName, selectedHierarchy.DimensionUniqueName);
        Assert.Equal("Calendar", selectedHierarchy.Caption);
        var selectedLevel = Assert.Single(schema.Levels);
        Assert.Equal(level, selectedLevel.UniqueName);
        Assert.Equal(hierarchy, selectedLevel.HierarchyUniqueName);
        Assert.Equal("Month", selectedLevel.Caption);
        // Atoti 6.2.2 reports LEVEL_CARDINALITY=1; actual members are verified independently below.
        Assert.Equal(1, selectedLevel.MemberCount);
        Assert.Equal(2, selectedLevel.Number);

        string[] expectedCaptions = Enumerable.Range(0, 24)
            .Select(offset => new DateOnly(2024, 1, 1).AddMonths(offset)
                .ToString("yyyy-MM", System.Globalization.CultureInfo.InvariantCulture))
            .ToArray();
        string[] expectedUniqueNames = expectedCaptions
            .Select(caption => $"{hierarchy}.[ALL].[AllMember].[{caption[..4]}].[{caption}]")
            .ToArray();
        var actualMembers = new List<Sbroenne.ExcelMcp.Core.Models.ExternalOlapMemberInfo>();
        string? continuationToken = null;
        for (int offset = 0; offset < expectedCaptions.Length; offset += pageSize)
        {
            var page = RequireSuccess(_connections.SearchOlapMembers(
                _fixture.BatchToken,
                connectionName,
                hierarchy,
                level,
                continuationToken: continuationToken,
                pageSize: pageSize));
            var expectedPageCaptions = expectedCaptions.Skip(offset).Take(pageSize).ToArray();
            var expectedPageUniqueNames = expectedUniqueNames.Skip(offset).Take(pageSize).ToArray();
            Assert.Equal(connectionName, page.ConnectionName);
            Assert.Equal(hierarchy, page.HierarchyUniqueName);
            Assert.Equal(level, page.LevelUniqueName);
            Assert.Null(page.SearchText);
            Assert.Equal(expectedPageCaptions.Length, page.ReturnedCount);
            Assert.Equal(1, page.TotalCount);
            Assert.Equal(24, page.ScannedCount);
            Assert.Equal(expectedPageCaptions, page.Members.Select(member => member.Caption));
            Assert.Equal(expectedPageUniqueNames, page.Members.Select(member => member.UniqueName));
            actualMembers.AddRange(page.Members);
            continuationToken = page.ContinuationToken;
            if (offset + expectedPageCaptions.Length < expectedCaptions.Length)
            {
                Assert.False(string.IsNullOrWhiteSpace(continuationToken));
            }
            else
            {
                Assert.Null(continuationToken);
            }
        }
        Assert.Equal(expectedCaptions, actualMembers.Select(member => member.Caption));
        Assert.Equal(expectedUniqueNames, actualMembers.Select(member => member.UniqueName));

        var november = RequireSuccess(_connections.SearchOlapMembers(
            _fixture.BatchToken, connectionName, hierarchy, level, searchText: "2025-11"));
        Assert.Equal("2025-11", november.SearchText);
        var currentMonth = Assert.Single(november.Members);
        Assert.Equal("2025-11", currentMonth.Caption);
        Assert.Equal($"{hierarchy}.[ALL].[AllMember].[2025].[2025-11]", currentMonth.UniqueName);
        Assert.Equal(1, november.ReturnedCount);
        Assert.Null(november.TotalCount);
        Assert.Null(november.ContinuationToken);
    }
}
