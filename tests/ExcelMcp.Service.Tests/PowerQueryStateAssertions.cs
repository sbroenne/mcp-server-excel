using System.Data.Common;
using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

internal static class PowerQueryStateAssertions
{
    internal static void AssertStored(
        PersistentServiceWorkbookTestScope scope,
        string name,
        string mCode,
        PowerQueryLoadMode mode,
        string? sheetName,
        string[] columns,
        object[][] rows)
    {
        var queries = scope.CreateCommands<IPowerQueryCommands>();
        var view = RequireSuccess(queries.View(scope.BatchToken, name));
        var listed = Assert.Single(RequireSuccess(queries.List(scope.BatchToken)).Queries,
            query => query.Name == name);
        var config = RequireSuccess(queries.GetLoadConfig(scope.BatchToken, name));
        var worksheet = mode is PowerQueryLoadMode.LoadToTable or PowerQueryLoadMode.LoadToBoth;
        var model = mode is PowerQueryLoadMode.LoadToDataModel or PowerQueryLoadMode.LoadToBoth;
        Assert.Equal(name, view.QueryName);
        Assert.Equal(mCode, view.MCode);
        Assert.Equal(mCode.Length, view.CharacterCount);
        Assert.Equal(mCode.Length > 80 ? mCode[..77] + "..." : mCode, listed.FormulaPreview);
        Assert.Equal(mCode.Length, listed.CharacterCount);
        Assert.Equal(mode, view.LoadMode);
        Assert.Equal(mode, listed.LoadMode);
        Assert.Equal(mode, config.LoadMode);
        Assert.Equal(sheetName, view.TargetSheet);
        Assert.Equal(sheetName, listed.TargetSheet);
        Assert.Equal(sheetName, config.TargetSheet);
        Assert.Equal(!worksheet && !model, view.IsConnectionOnly);
        Assert.Equal(!worksheet && !model, listed.IsConnectionOnly);
        Assert.Equal(model, view.IsLoadedToDataModel);
        Assert.Equal(model, listed.IsLoadedToDataModel);
        Assert.Equal(model, config.IsLoadedToDataModel);
        Assert.Equal(worksheet || model, view.HasConnection);
        Assert.Equal(worksheet || model, config.HasConnection);
        AssertNativeIdentity(scope, name, mCode, worksheet, model);
        if (worksheet)
        {
            Assert.NotNull(sheetName);
            var range = scope.CreateCommands<IRangeCommands>();
            var values = RequireSuccess(range.GetValues(scope.BatchToken, sheetName,
                $"A1:{ColumnName(columns.Length)}{rows.Length + 1}")).Values;
            Assert.Equal(rows.Length + 1, values.Count);
            Assert.Equal(columns, values[0].Select(value => Convert.ToString(value, CultureInfo.InvariantCulture)));
            AssertRows(rows, values.Skip(1).ToList());
        }
        if (model)
        {
            var dataModel = scope.CreateCommands<IDataModelCommands>();
            var result = RequireSuccess(dataModel.Evaluate(scope.BatchToken,
                $"EVALUATE '{name.Replace("'", "''", StringComparison.Ordinal)}'"));
            Assert.Equal(columns.Length, result.ColumnCount);
            Assert.Equal(rows.Length, result.RowCount);
            Assert.Equal(columns.Select(column => $"{name}[{column}]"), result.Columns);
            AssertRows(rows, result.Rows);
        }
    }

    internal static void AssertRows(object[][] expected, List<List<object?>> actual)
    {
        Assert.Equal(expected.Length, actual.Count);
        for (var row = 0; row < expected.Length; row++)
        {
            Assert.Equal(expected[row].Length, actual[row].Count);
            for (var column = 0; column < expected[row].Length; column++)
            {
                Assert.NotNull(actual[row][column]);
                if (expected[row][column] is string text)
                {
                    Assert.Equal(text, Assert.IsType<string>(actual[row][column]));
                }
                else if (expected[row][column] is bool flag)
                {
                    Assert.Equal(flag, Assert.IsType<bool>(actual[row][column]));
                }
                else
                {
                    Assert.Equal(Convert.ToDecimal(expected[row][column], CultureInfo.InvariantCulture),
                        Convert.ToDecimal(actual[row][column], CultureInfo.InvariantCulture));
                }
            }
        }
    }

    internal static void AssertRemoved(PersistentServiceWorkbookTestScope scope, string name)
    {
        var queries = scope.CreateCommands<IPowerQueryCommands>();
        Assert.DoesNotContain(RequireSuccess(queries.List(scope.BatchToken)).Queries,
            query => string.Equals(query.Name, name, StringComparison.OrdinalIgnoreCase));
        AssertNativeIdentity(scope, name, null, false, false);
    }

    private static void AssertNativeIdentity(
        PersistentServiceWorkbookTestScope scope, string name, string? mCode,
        bool worksheet, bool inModel) =>
        scope.ExecuteRawVerification((context, _) =>
        {
            Excel.Queries? queries = null;
            Excel.WorkbookQuery? query = null;
            Excel.Connections? connections = null;
            Excel.Model? model = null;
            Excel.ModelTables? tables = null;
            try
            {
                queries = context.Book.Queries;
                if (mCode is not null)
                {
                    query = queries.Item(name);
                    Assert.Equal(name, query.Name);
                    Assert.Equal(mCode, query.Formula);
                }
                else
                {
                    for (var index = 1; index <= queries.Count; index++)
                    {
                        Excel.WorkbookQuery? candidate = null;
                        try
                        {
                            candidate = queries.Item(index);
                            Assert.False(string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase));
                        }
                        finally { ComUtilities.Release(ref candidate); }
                    }
                }
                connections = context.Book.Connections;
                var matches = 0;
                var modelConnections = 0;
                var modelConnectionNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (var index = 1; index <= connections.Count; index++)
                {
                    Excel.WorkbookConnection? connection = null;
                    Excel.OLEDBConnection? oledb = null;
                    try
                    {
                        connection = connections.Item(index);
                        if (connection.Type != Excel.XlConnectionType.xlConnectionTypeOLEDB) { continue; }
                        oledb = connection.OLEDBConnection;
                        string text = Convert.ToString(oledb.Connection, CultureInfo.InvariantCulture) ?? "";
                        if (text.StartsWith("OLEDB;", StringComparison.OrdinalIgnoreCase)) { text = text[6..]; }
                        var properties = new DbConnectionStringBuilder { ConnectionString = text };
                        if (!properties.TryGetValue("Location", out var location) ||
                            !string.Equals(Convert.ToString(location, CultureInfo.InvariantCulture), name,
                                StringComparison.OrdinalIgnoreCase)) { continue; }
                        matches++;
                        if (connection.InModel)
                        {
                            modelConnections++;
                            Assert.True(modelConnectionNames.Add(connection.Name));
                        }
                        Assert.False(oledb.Refreshing);
                    }
                    finally
                    {
                        ComUtilities.Release(ref oledb);
                        ComUtilities.Release(ref connection);
                    }
                }
                Assert.Equal((worksheet ? 1 : 0) + (inModel ? 1 : 0), matches);
                Assert.Equal(inModel ? 1 : 0, modelConnections);
                model = context.Book.Model;
                tables = model.ModelTables;
                var tableMatches = 0;
                for (var index = 1; index <= tables.Count; index++)
                {
                    Excel.ModelTable? table = null;
                    Excel.WorkbookConnection? source = null;
                    try
                    {
                        table = tables.Item(index);
                        if (string.Equals(table.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            tableMatches++;
                            source = table.SourceWorkbookConnection;
                            Assert.True(source.InModel);
                            Assert.Contains(source.Name, modelConnectionNames);
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref source);
                        ComUtilities.Release(ref table);
                    }
                }
                Assert.Equal(inModel ? 1 : 0, tableMatches);
            }
            finally
            {
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref model);
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
            }
        });

    private static string ColumnName(int count)
    {
        var name = "";
        while (count > 0)
        {
            count--;
            name = (char)('A' + count % 26) + name;
            count /= 26;
        }
        return name;
    }

    private static T RequireSuccess<T>(T result) where T : ResultBase
    {
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage), result.ErrorMessage);
        return result;
    }
}
