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
        Assert.True(allSchema.Hierarchies.Count > 1, "The test cube must expose unused hierarchies.");
        Assert.All(allSchema.Hierarchies, hierarchy =>
            Assert.Contains(allSchema.Dimensions, dimension => dimension.UniqueName == hierarchy.DimensionUniqueName));
        var schema = RequireSuccess(_connections.DiscoverOlapSchema(
            _fixture.BatchToken,
            connectionName,
            hierarchyUniqueName));
        Assert.Contains(schema.Hierarchies, hierarchy => hierarchy.UniqueName == hierarchyUniqueName);
        Assert.Contains(schema.Levels, level => level.UniqueName == levelUniqueName);
        Assert.Equal(hierarchyUniqueName, Assert.Single(schema.Hierarchies).UniqueName);
        Assert.All(schema.Levels, level => Assert.Equal(hierarchyUniqueName, level.HierarchyUniqueName));
        var filteredLevel = RequireSuccess(_connections.DiscoverOlapSchema(
            _fixture.BatchToken, connectionName, hierarchyUniqueName, levelUniqueName));
        Assert.Equal(levelUniqueName, Assert.Single(filteredLevel.Levels).UniqueName);

        var firstPage = RequireSuccess(_connections.SearchOlapMembers(
            _fixture.BatchToken,
            connectionName,
            hierarchyUniqueName,
            levelUniqueName,
            pageSize: 1));
        var firstMember = Assert.Single(firstPage.Members);
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
        Assert.NotEqual(firstMember.UniqueName, secondMember.UniqueName);
        Assert.True(string.CompareOrdinal(firstMember.UniqueName, secondMember.UniqueName) < 0);
        Assert.InRange(firstPage.ScannedCount, 1, 10000);
        Assert.InRange(secondPage.ScannedCount, 1, 10000);
        if (firstPage.TotalCount.HasValue)
        {
            Assert.Equal(firstPage.TotalCount - firstPage.ReturnedCount, firstPage.OmittedCount);
        }
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
        Assert.Equal($"{hierarchy}.[ALL].[AllMember].[Current Month]", Assert.Single(current.Members).UniqueName);
        Assert.Null(current.TotalCount);
        Assert.Null(current.OmittedCount);
    }
}
