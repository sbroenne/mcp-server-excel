using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePowerQueryMutationFailureTests : IAsyncLifetime, IDisposable
{
    private const string ValidMCode = "#table({\"X\"}, {{1}})";
    private readonly PersistentServiceWorkbookFixture _fixture = new();

    public Task InitializeAsync() => _fixture.InitializeAsync();
    public Task DisposeAsync() => _fixture.DisposeAsync();
    public void Dispose() => _fixture.Dispose();

    [Theory]
    [InlineData("DataSource.Error", "This server does not support Windows credentials.", "Authentication", "Organizational account")]
    [InlineData("DataSource.Error", "Analysis Services does not support the 'en-DE' culture", "Culture", "Culture = \"en-US\"")]
    public async Task LoadTo_RecognizedFailure_PreservesDefinitionAndReportsActualPartialLoad(
        string reason, string message, string category, string hint)
    {
        const string queryName = "FailedLoad";
        var code = FirewallMCode(reason, message);
        await _fixture.SendAsync("powerquery.create",
            new { queryName, mCode = code, loadDestination = "connection-only" });

        var response = await _fixture.SendForFailureAsync("powerquery.load-to",
            new { queryName, loadDestination = "worksheet", targetSheet = queryName });
        Assert.Equal(category, response.ErrorCategory);
        Assert.Contains(message, response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains(hint, response.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal("0x800A03EC", response.HResult);
        var state = ReadPartialState(response);
        Assert.False(state.TryGetProperty("mCodeUpdated", out _));
        Assert.False(state.TryGetProperty("queryCreated", out _));
        AssertState(state, queryName, "load-to-table", queryName);
        AssertStoredDefinition(queryName, code, PowerQueryLoadMode.LoadToTable, queryName);
        await _fixture.SendAsync("powerquery.update", new { queryName, mCode = ValidMCode });
        AssertWorksheetValue(queryName, 1);
    }

    [Theory]
    [InlineData("Formula.Firewall", "Query references other queries or steps, so it may not directly access a data source.")]
    [InlineData("Privacy", "Information is needed in order to combine data. Please specify a privacy level for each data source.")]
    public async Task Create_PrivacyLoadFailure_ReportsSurvivingWorksheetNotRequestedBoth(
        string reason, string message)
    {
        const string queryName = "FailedCreate";
        var mCode = FirewallMCode(reason, message);

        var response = await _fixture.SendForFailureAsync("powerquery.create",
            new { queryName, mCode, loadDestination = "both" });

        AssertStoredDefinition(queryName, mCode, PowerQueryLoadMode.LoadToTable, queryName);
        Assert.Equal(reason == "Formula.Firewall" ? "PrivacyFirewall" : "Privacy", response.ErrorCategory);
        Assert.Equal("PowerQueryCommandException", response.ExceptionType);
        Assert.Equal("powerquery.create", response.Command);
        Assert.Equal("0x800A03EC", response.HResult);
        Assert.Contains(message, response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains(queryName, response.ErrorMessage, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(response.InnerError));
        var state = ReadPartialState(response);
        Assert.True(state.GetProperty("queryCreated").GetBoolean());
        Assert.False(state.TryGetProperty("mCodeUpdated", out _));
        AssertState(state, queryName, "load-to-table", queryName);

        await _fixture.SendAsync("powerquery.update", new { queryName, mCode = "#table({\"X\"}, {{2}})" });
        AssertWorksheetValue(queryName, 2);
    }

    [Theory]
    [InlineData("Formula.Firewall", "Query references other queries or steps, so it may not directly access a data source.", "PrivacyFirewall")]
    [InlineData("Privacy", "Information is needed in order to combine data. Please specify a privacy level for each data source.", "Privacy")]
    [InlineData("DataSource.Error", "Analysis Services does not support the 'en-DE' culture", "Culture")]
    public async Task Update_RecognizedRefreshFailure_ReportsSavedCodeAndPreservesOldRows(
        string reason, string message, string category)
    {
        const string queryName = "FailedUpdate";
        await _fixture.SendAsync("powerquery.create", new { queryName, mCode = ValidMCode });
        AssertWorksheetValue(queryName, 1);
        var mCode = FirewallMCode(reason, message);

        var response = await _fixture.SendForFailureAsync("powerquery.update", new { queryName, mCode });

        AssertStoredDefinition(queryName, mCode, PowerQueryLoadMode.LoadToTable, queryName);
        AssertWorksheetValue(queryName, 1);
        Assert.Equal(category, response.ErrorCategory);
        Assert.Equal("PowerQueryCommandException", response.ExceptionType);
        Assert.Equal("powerquery.update", response.Command);
        Assert.Equal("0x800A03EC", response.HResult);
        Assert.Contains(message, response.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains(queryName, response.ErrorMessage, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(response.InnerError));
        Assert.DoesNotContain("M code was not changed", response.ErrorMessage, StringComparison.Ordinal);
        var state = ReadPartialState(response);
        Assert.True(state.GetProperty("mCodeUpdated").GetBoolean());
        Assert.False(state.TryGetProperty("queryCreated", out _));
        AssertState(state, queryName, "load-to-table", queryName);

        await _fixture.SendAsync("powerquery.update", new { queryName, mCode = "#table({\"X\"}, {{2}})" });
        AssertWorksheetValue(queryName, 2);
    }

    [Theory]
    [InlineData("Guard")]
    [InlineData("privacy level")]
    public async Task Create_OccupiedDestination_ReportsConnectionOnlyDefinitionAndPreservesCell(string sheetName)
    {
        const string queryName = "OccupiedCreate";
        await _fixture.SendAsync("sheet.create", new { sheetName });
        await _fixture.SendAsync("range.set-values",
            new { sheetName, rangeAddress = "A1", values = new[] { new object[] { "keep" } } });

        var response = await _fixture.SendForFailureAsync("powerquery.create",
            new { queryName, mCode = ValidMCode, targetSheet = sheetName });

        AssertStoredDefinition(queryName, ValidMCode, PowerQueryLoadMode.ConnectionOnly, null);
        var ranges = _fixture.CreateCommands<Sbroenne.ExcelMcp.Core.Commands.Range.IRangeCommands>();
        var values = ranges.GetValues(_fixture.BatchToken, sheetName, "A1");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal("keep", Assert.Single(Assert.Single(values.Values)));
        Assert.NotEqual("Privacy", response.ErrorCategory);
        Assert.Contains("already contains data", response.ErrorMessage, StringComparison.Ordinal);
        var state = ReadPartialState(response);
        Assert.True(state.GetProperty("queryCreated").GetBoolean());
        AssertState(state, queryName, "connection-only", null);
    }

    [Fact]
    public async Task Create_ExistingQuery_OmitsPartialStateAndPreservesDefinition()
    {
        const string queryName = "ExistingQuery";
        await _fixture.SendAsync("powerquery.create",
            new { queryName, mCode = ValidMCode, loadDestination = "connection-only" });

        var response = await _fixture.SendForFailureAsync("powerquery.create",
            new { queryName, mCode = "#table({\"X\"}, {{99}})" });

        AssertStoredDefinition(queryName, ValidMCode, PowerQueryLoadMode.ConnectionOnly, null);
        Assert.Contains("already exists", response.ErrorMessage, StringComparison.Ordinal);
        Assert.False(JsonSerializer.SerializeToElement(response, ServiceProtocol.JsonOptions)
            .TryGetProperty("partialState", out _));
    }

    [Theory]
    [InlineData(PowerQueryLoadMode.LoadToDataModel, "data-model", "load-to-data-model")]
    [InlineData(PowerQueryLoadMode.LoadToBoth, "both", "load-to-both")]
    public async Task Update_PrivacyFailure_PreservesModelLoadAndOldData(
        PowerQueryLoadMode mode, string destination, string serializedMode)
    {
        const string queryName = "RetainedModel";
        const string seed = "Table.TransformColumnTypes(#table({\"X\"}, {{1}}), {{\"X\", Int64.Type}})";
        var mCode = FirewallMCode("Formula.Firewall", "Please specify a privacy level for each data source.");
        await using var scope = _fixture.CreateScope();
        await _fixture.SendAsync("powerquery.create",
            new { queryName, mCode = seed, loadDestination = destination });
        var sheet = mode == PowerQueryLoadMode.LoadToBoth ? queryName : null;
        PowerQueryStateAssertions.AssertStored(scope, queryName, seed, mode, sheet, ["X"], [[1]]);

        var response = await _fixture.SendForFailureAsync("powerquery.update", new { queryName, mCode });

        PowerQueryStateAssertions.AssertStored(scope, queryName, mCode, mode, sheet, ["X"], [[1]]);
        Assert.Equal("PrivacyFirewall", response.ErrorCategory);
        Assert.Equal("PowerQueryCommandException", response.ExceptionType);
        var state = ReadPartialState(response);
        Assert.Equal(queryName, state.GetProperty("queryName").GetString());
        Assert.True(state.GetProperty("mCodeUpdated").GetBoolean());
        Assert.True(state.GetProperty("queryExists").GetBoolean());
        Assert.Equal(serializedMode, state.GetProperty("loadMode").GetString());
        Assert.True(state.GetProperty("isLoadedToDataModel").GetBoolean());
        if (sheet is null)
            Assert.False(state.TryGetProperty("targetSheet", out _));
        else
            Assert.Equal(sheet, state.GetProperty("targetSheet").GetString());
    }

    private static string FirewallMCode(string reason, string message) =>
        $"error Error.Record(\"{reason}\", \"{message}\", null)";

    private static JsonElement ReadPartialState(ServiceResponse response) =>
        JsonSerializer.Deserialize<JsonElement>(Assert.IsType<string>(response.PartialState));

    private static void AssertState(JsonElement state, string name, string mode, string? sheet)
    {
        Assert.Equal(name, state.GetProperty("queryName").GetString());
        Assert.True(state.GetProperty("queryExists").GetBoolean());
        Assert.Equal(mode, state.GetProperty("loadMode").GetString());
        Assert.False(state.GetProperty("isLoadedToDataModel").GetBoolean());
        Assert.False(state.TryGetProperty("stateInspectionError", out _));
        if (sheet is null)
            Assert.False(state.TryGetProperty("targetSheet", out _));
        else
            Assert.Equal(sheet, state.GetProperty("targetSheet").GetString());
    }

    private void AssertStoredDefinition(string name, string mCode, PowerQueryLoadMode mode, string? sheet)
    {
        var queries = _fixture.CreateCommands<IPowerQueryCommands>();
        var view = queries.View(_fixture.BatchToken, name);
        Assert.True(view.Success, view.ErrorMessage);
        Assert.Equal(mCode, view.MCode);
        Assert.Equal(mode, view.LoadMode);
        Assert.Equal(sheet, view.TargetSheet);
        Assert.False(view.IsLoadedToDataModel);
        var list = queries.List(_fixture.BatchToken);
        Assert.True(list.Success, list.ErrorMessage);
        Assert.Equal(mode, Assert.Single(list.Queries, query => query.Name == name).LoadMode);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Queries? collection = null;
            Excel.WorkbookQuery? query = null;
            Excel.Model? model = null;
            Excel.ModelTables? tables = null;
            try
            {
                collection = context.Book.Queries;
                query = collection.Item(name);
                Assert.Equal(name, query.Name);
                Assert.Equal(mCode, query.Formula);
                model = context.Book.Model;
                tables = model.ModelTables;
                Assert.Equal(0, tables.Count);
            }
            finally
            {
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref model);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref collection);
            }
        });
    }

    private void AssertWorksheetValue(string sheet, int expected)
    {
        var ranges = _fixture.CreateCommands<Sbroenne.ExcelMcp.Core.Commands.Range.IRangeCommands>();
        var values = ranges.GetValues(_fixture.BatchToken, sheet, "A1:A2");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal(2, values.Values.Count);
        Assert.Equal("X", Assert.Single(values.Values[0]));
        Assert.Equal(expected, Convert.ToInt32(Assert.Single(values.Values[1]), CultureInfo.InvariantCulture));
    }
}
