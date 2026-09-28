using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePowerQueryEvaluateTests
{
    [Fact]
    public void Evaluate_InvalidMCode_ThrowsError()
    {
        const string invalidMCode = """
            let
                Source = UndefinedFunction()
            in
                Source
            """;
        var initialState = EvaluateObjectCounts();

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _queries.Evaluate(_fixture.BatchToken, invalidMCode));

        Assert.Contains(
            "Expression.Error",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        Assert.Contains(
            "PowerQueryCommandException",
            exception.Message,
            StringComparison.Ordinal);
        Assert.Contains(
            "Expression",
            exception.Message,
            StringComparison.Ordinal);
        Assert.Equal(initialState, EvaluateObjectCounts());
    }

    [Fact]
    public void Evaluate_AfterExecution_CleansUpTempObjects()
    {
        const string mCode = """
            let
                Source = #table({"X"}, {{1}})
            in
                Source
            """;
        var initialQueries = _queries.List(_fixture.BatchToken);
        var initialQueryCount = initialQueries.Queries.Count;
        var initialState = EvaluateObjectCounts();

        var result = _queries.Evaluate(_fixture.BatchToken, mCode);

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        var finalQueries = _queries.List(_fixture.BatchToken);
        Assert.Equal(initialQueryCount, finalQueries.Queries.Count);
        Assert.Equal(initialState, EvaluateObjectCounts());
        Assert.Equal("X", Assert.Single(result.Columns));
        Assert.Equal(
            1.0,
            Convert.ToDouble(
                Assert.Single(Assert.Single(result.Rows)),
                System.Globalization.CultureInfo.InvariantCulture));
    }

    private (int Sheets, int Queries, int Connections) EvaluateObjectCounts() =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Queries? queries = null;
            Excel.Connections? connections = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                queries = ctx.Book.Queries;
                connections = ctx.Book.Connections;
                return (sheets.Count, queries.Count, connections.Count);
            }
            finally
            {
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref queries);
                ComUtilities.Release(ref sheets);
            }
        });
}
