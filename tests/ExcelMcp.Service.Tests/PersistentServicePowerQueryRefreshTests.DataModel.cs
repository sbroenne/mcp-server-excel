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
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            ValidMCode,
            PowerQueryLoadMode.LoadToDataModel);
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _queries.Update(
            _fixture.BatchToken,
            queryName,
            invalidMCode,
            refresh: false);

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
        _queries.Create(
            _fixture.BatchToken,
            queryName,
            mCode,
            PowerQueryLoadMode.LoadToDataModel);
        _fixture.RegisterPowerQueryForCleanup(queryName);

        var result = _queries.Refresh(
            _fixture.BatchToken,
            queryName,
            TimeSpan.FromMinutes(1));

        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
        Assert.False(result.HasErrors);
    }
}
