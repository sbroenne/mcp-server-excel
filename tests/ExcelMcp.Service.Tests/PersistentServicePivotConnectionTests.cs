using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePivotConnectionTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public async Task SetConnection_ExternalPivotUsesTargetAndPreservesLayoutAndOtherPivot()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var prefix = $"External_{Guid.NewGuid():N}";
        var oldConnection = prefix + "_old";
        var targetConnection = prefix + "_target";
        var selected = prefix + "_selected";
        var other = prefix + "_other";
        var oldFile = _fixture.CreateInputFile(".csv", "Region,Sales\r\nNorth,10\r\nSouth,20\r\n");
        var targetDirectory = Path.Combine(Path.GetDirectoryName(oldFile)!, prefix);
        Directory.CreateDirectory(targetDirectory);
        var targetFile = Path.Combine(targetDirectory, Path.GetFileName(oldFile));
        File.WriteAllText(targetFile, "Region,Sales\r\nNorth,15\r\nSouth,25\r\n");
        CreateExternalConnection(oldConnection, oldFile);
        CreateExternalConnection(targetConnection, targetFile);
        CreateExternalPivot(sheet, "A1", selected, oldConnection);
        CreateExternalPivot(sheet, "G1", other, oldConnection);
        var commands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        foreach (var name in new[] { selected, other })
        {
            RequireSuccess(commands.AddRowField(_fixture.BatchToken, name, "Region"));
            RequireSuccess(commands.AddValueField(_fixture.BatchToken, name, "Sales"));
        }
        RequireSuccess(commands.SetFieldFilter(_fixture.BatchToken, selected, "Region", ["North"]));
        RequireSuccess(commands.SetLayoutOptions(_fixture.BatchToken, selected,
            new PivotLayoutOptions { RowLayout = 1, StyleName = "PivotStyleMedium9", PreserveFormatting = true }));
        var beforeOther = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = other });
        Assert.True(beforeOther.Success, beforeOther.ErrorMessage);
        var beforeOtherSource = _fixture.Send("pivottable.get-connection",
            new { sheetName = sheet, pivotTableName = other });
        var change = _fixture.Send("pivottable.set-connection",
            new { sheetName = sheet, pivotTableName = selected, connectionName = targetConnection });
        Assert.True(change.Success, change.ErrorMessage);
        using var changed = JsonDocument.Parse(change.Result!);
        Assert.Equal(sheet, changed.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal(selected, changed.RootElement.GetProperty("pivotTableName").GetString());
        Assert.Equal(targetConnection, changed.RootElement.GetProperty("connectionName").GetString());
        RequireSuccess(commands.Refresh(_fixture.BatchToken, selected));
        var read = RequireSuccess(commands.Read(_fixture.BatchToken, selected));
        Assert.Contains(read.Fields, field => field.Name == "Region");
        Assert.Contains(read.Fields, field => field.Name == "Sales");
        var layout = RequireSuccess(commands.GetLayoutOptions(_fixture.BatchToken, selected));
        Assert.Equal("PivotStyleMedium9", layout.StyleName);
        Assert.True(layout.PreserveFormatting);
        Assert.Equal(1, Assert.Single(layout.RowFields).RowLayout);
        var data = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = selected });
        using var values = JsonDocument.Parse(data.Result!);
        Assert.Equal(15, values.RootElement.GetProperty("values")[1][1].GetDouble());
        Assert.Equal("North", values.RootElement.GetProperty("values")[1][0].GetString());
        Assert.Equal(15, values.RootElement.GetProperty("values")[2][1].GetDouble());
        Assert.Equal(3, values.RootElement.GetProperty("values").GetArrayLength());
        var otherSource = _fixture.Send("pivottable.get-connection", new { sheetName = sheet, pivotTableName = other });
        using var beforeSource = JsonDocument.Parse(beforeOtherSource.Result!);
        using var afterSource = JsonDocument.Parse(otherSource.Result!);
        Assert.Equal(beforeSource.RootElement.GetProperty("connectionName").GetString(),
            afterSource.RootElement.GetProperty("connectionName").GetString());
        Assert.Equal(other, afterSource.RootElement.GetProperty("pivotTableName").GetString());
        Assert.Equal(sheet, afterSource.RootElement.GetProperty("sheetName").GetString());
        Assert.Single(afterSource.RootElement.GetProperty("sharedPivotTables").EnumerateArray());
        var afterOther = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = other });
        Assert.Equal(beforeOther.Result, afterOther.Result);
        await _fixture.SaveAndReopenAsync();
        var persisted = _fixture.Send("pivottable.get-connection",
            new { sheetName = sheet, pivotTableName = selected });
        using var persistedSource = JsonDocument.Parse(persisted.Result!);
        Assert.Equal(targetConnection, persistedSource.RootElement.GetProperty("connectionName").GetString());
        var persistedData = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = selected });
        Assert.Equal(data.Result, persistedData.Result);
        Assert.Equal("PivotStyleMedium9", RequireSuccess(commands.GetLayoutOptions(_fixture.BatchToken, selected)).StyleName);
    }

    [Theory]
    [InlineData("missing", "not found")]
    [InlineData("protected", "Unprotect")]
    [InlineData("slicer", "slicer")]
    [InlineData("shared", "share")]
    public async Task SetConnection_RejectsBeforeChangingExternalPivot(string scenario, string expectedError)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var prefix = $"Guard_{Guid.NewGuid():N}";
        var oldConnection = prefix + "_old";
        var targetConnection = prefix + "_target";
        var name = prefix + "_pivot";
        var file = _fixture.CreateInputFile(".csv", "Region,Sales\r\nNorth,10\r\nSouth,20\r\n");
        CreateExternalConnection(oldConnection, file);
        CreateExternalConnection(targetConnection, file);
        CreateExternalPivot(sheet, "A1", name, oldConnection);
        var commands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(commands.AddRowField(_fixture.BatchToken, name, "Region"));
        RequireSuccess(commands.AddValueField(_fixture.BatchToken, name, "Sales"));
        if (scenario == "shared")
            CreateExternalPivot(sheet, "G1", prefix + "_other", oldConnection, name);
        if (scenario == "slicer")
            RequireSuccess(commands.CreateSlicer(_fixture.BatchToken, name, "Region", prefix + "_slicer", sheet, "J1"));
        if (scenario == "protected")
        {
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Worksheet? worksheet = null;
                try
                {
                    worksheet = ComUtilities.FindSheet(context.Book, sheet);
                    worksheet.Protect();
                }
                finally
                {
                    ComUtilities.Release(ref worksheet);
                }
            });
        }
        var before = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        var failure = await _fixture.SendForFailureAsync("pivottable.set-connection", new
        {
            sheetName = sheet,
            pivotTableName = name,
            connectionName = scenario == "missing" ? "MissingConnection" : targetConnection
        });
        Assert.Contains(expectedError, failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var after = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        Assert.Equal(before.Result, after.Result);
        var actual = _fixture.Send("pivottable.get-connection", new { sheetName = sheet, pivotTableName = name });
        using var source = JsonDocument.Parse(actual.Result!);
        Assert.Equal(oldConnection, source.RootElement.GetProperty("connectionName").GetString());
    }

    [Fact]
    public async Task ConnectionActions_WorksheetPivotReportsSourceAndRejectsChangeWithoutMutation()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Connection_{Guid.NewGuid():N}";
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B3",
            [["Region", "Sales"], ["North", 10], ["South", 20]]));
        var pivot = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivot.CreateFromRange(_fixture.BatchToken, sheet, "A1:B3", sheet, "D1", name));
        RequireSuccess(pivot.AddRowField(_fixture.BatchToken, name, "Region"));
        RequireSuccess(pivot.AddValueField(_fixture.BatchToken, name, "Sales"));
        var before = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        Assert.True(before.Success, before.ErrorMessage);

        var read = _fixture.Send("pivottable.get-connection", new { sheetName = sheet, pivotTableName = name });
        Assert.True(read.Success, read.ErrorMessage);
        using var source = JsonDocument.Parse(read.Result!);
        Assert.Equal(sheet, source.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal(name, source.RootElement.GetProperty("pivotTableName").GetString());
        Assert.False(source.RootElement.GetProperty("isExternal").GetBoolean());
        Assert.False(source.RootElement.GetProperty("isOlap").GetBoolean());
        Assert.True(source.RootElement.GetProperty("cacheIndex").GetInt32() > 0);
        Assert.False(source.RootElement.TryGetProperty("connectionString", out _));
        var shared = Assert.Single(source.RootElement.GetProperty("sharedPivotTables").EnumerateArray());
        Assert.Equal(sheet, shared.GetProperty("sheetName").GetString());
        Assert.Equal(name, shared.GetProperty("pivotTableName").GetString());

        var rejected = await _fixture.SendForFailureAsync("pivottable.set-connection",
            new { sheetName = sheet, pivotTableName = name, connectionName = "MissingConnection" });
        Assert.False(rejected.Success);
        Assert.Contains("external", rejected.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var after = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(before.Result, after.Result);
        var unchanged = _fixture.Send("pivottable.get-connection", new { sheetName = sheet, pivotTableName = name });
        Assert.True(unchanged.Success, unchanged.ErrorMessage);
        Assert.Equal(read.Result, unchanged.Result);
    }

    [Fact]
    public async Task GetConnection_WrongWorksheetDoesNotSelectPivotFromAnotherSheet()
    {
        var sourceSheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var wrongSheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Selection_{Guid.NewGuid():N}";
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sourceSheet, "A1:B2",
            [["Region", "Sales"], ["North", 10]]));
        var pivot = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(pivot.CreateFromRange(_fixture.BatchToken, sourceSheet, "A1:B2", sourceSheet, "D1", name));

        var failure = await _fixture.SendForFailureAsync("pivottable.get-connection",
            new { sheetName = wrongSheet, pivotTableName = name });
        Assert.False(failure.Success);
        Assert.Contains("not found", failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var actual = _fixture.Send("pivottable.get-connection",
            new { sheetName = sourceSheet, pivotTableName = name });
        Assert.True(actual.Success, actual.ErrorMessage);
        using var result = JsonDocument.Parse(actual.Result!);
        Assert.Equal(sourceSheet, result.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal(name, result.RootElement.GetProperty("pivotTableName").GetString());
    }

    private void CreateExternalConnection(string name, string file)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.OLEDBConnection? oleDb = null;
            try
            {
                connections = context.Book.Connections;
                connection = connections.Add2(name, "Synthetic external PivotTable fixture",
                    $"OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source={Path.GetDirectoryName(file)};Extended Properties=\"Text;HDR=Yes;FMT=Delimited\";",
                    $"SELECT Region, Sales FROM [{Path.GetFileName(file)}]", Excel.XlCmdType.xlCmdSql);
                oleDb = connection.OLEDBConnection;
                oleDb.BackgroundQuery = false;
                Assert.Equal(name, connection.Name);
            }
            finally
            {
                ComUtilities.Release(ref oleDb);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });
    }

    private void CreateExternalPivot(string sheetName, string address, string name, string connectionName,
        string? sharedPivotName = null)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.PivotCaches? caches = null;
            Excel.PivotCache? cache = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? destination = null;
            Excel.PivotTable? pivot = null;
            Excel.PivotTables? pivots = null;
            Excel.PivotTable? original = null;
            try
            {
                connections = context.Book.Connections;
                connection = connections.Item(connectionName);
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                if (sharedPivotName is null)
                {
                    caches = context.Book.PivotCaches();
                    cache = caches.Create(Excel.XlPivotTableSourceType.xlExternal, connection);
                }
                else
                {
                    pivots = sheet.PivotTables();
                    original = pivots.Item(sharedPivotName);
                    cache = original.PivotCache();
                }
                destination = sheet.Range[address];
                pivot = cache.CreatePivotTable(destination, name);
                Assert.Equal(name, pivot.Name);
                Assert.Equal(Excel.XlPivotTableSourceType.xlExternal, cache.SourceType);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
                ComUtilities.Release(ref original);
                ComUtilities.Release(ref pivots);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });
    }
}
