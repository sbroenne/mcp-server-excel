using System.Globalization;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryRefreshTests
{
    [Fact]
    public void Refresh_DataModelQueryWithInvalidMCode_ThrowsError()
    {
        var queryName = UniqueName("BrokenDataModel");
        const string invalidMCode =
            "let Source = UndefinedReference in Source";
        var guard = CreateRefreshGuard();
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            ValidMCode,
            PowerQueryLoadMode.LoadToDataModel));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _storedSources.Add(queryName, ValidMCode);
        AssertModelValue(queryName, 1);
        StageSource(queryName, invalidMCode);
        AssertModelValue(queryName, 1);
        var before = SnapshotQueries();

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Refresh(
                _fixture.BatchToken,
                queryName,
                TimeSpan.FromMinutes(1)));

        Assert.True(
            exception.Message.Contains(
                "Expression.Error",
                StringComparison.OrdinalIgnoreCase)
            || exception.Message.Contains(
                "Data Model",
                StringComparison.OrdinalIgnoreCase)
            || exception.Message.Contains(
                "couldn't get data",
                StringComparison.OrdinalIgnoreCase),
            $"Expected Power Query/Data Model error but got: {exception.Message}");
        AssertModelValue(queryName, 1);
        Assert.Equal(before, SnapshotQueries());
        AssertRefreshGuard(guard);
    }

    [Fact]
    public void Refresh_ValidDataModelQuery_Succeeds()
    {
        var queryName = UniqueName("ValidDataModel");
        const string mCode = """
            let
                Source = #table(
                    {"Category", "Amount"},
                    {{"Sales", 1000}, {"Marketing", 500}})
            in
                Source
            """;
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            mCode,
            PowerQueryLoadMode.LoadToDataModel));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _storedSources.Add(queryName, mCode);
        AssertCategoryAmounts(queryName, 1000, 500);
        StageSource(queryName,
            """
            let
                Source = #table(
                    {"Category", "Amount"},
                    {{"Sales", 1500}, {"Marketing", 750}})
            in
                Source
            """);
        AssertCategoryAmounts(queryName, 1000, 500);

        var result = RequireSuccess(_queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromMinutes(1)));

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors);
        AssertRefreshMetadata(result, queryName, null);
        AssertCategoryAmounts(queryName, 1500, 750);
    }

    private void AssertModelValue(string queryName, int expected)
    {
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, _storedSources[queryName],
            PowerQueryLoadMode.LoadToDataModel, null, ["X"], [[expected]]);
    }

    private void AssertCategoryAmounts(string queryName, int sales, int marketing)
    {
        var model = _fixture.CreateCommands<IDataModelCommands>();
        var result = RequireSuccess(model.Evaluate(_fixture.BatchToken,
            $"EVALUATE '{queryName}' ORDER BY '{queryName}'[Category]"));
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, _storedSources[queryName],
            PowerQueryLoadMode.LoadToDataModel, null, ["Category", "Amount"],
            [["Sales", sales], ["Marketing", marketing]]);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(2, result.RowCount);
        Assert.Collection(result.Rows,
            row =>
            {
                Assert.Equal(2, row.Count);
                Assert.Equal("Marketing", row[0]);
                Assert.Equal(marketing, Convert.ToInt32(row[1], CultureInfo.InvariantCulture));
            },
            row =>
            {
                Assert.Equal(2, row.Count);
                Assert.Equal("Sales", row[0]);
                Assert.Equal(sales, Convert.ToInt32(row[1], CultureInfo.InvariantCulture));
            });
    }
}
