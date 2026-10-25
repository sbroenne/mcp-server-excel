using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    private static readonly int[] s_legacyFixedWidths = [8];
    private static readonly int[] s_legacyFixedTypes =
        [(int)Excel.XlColumnDataType.xlTextFormat, (int)Excel.XlColumnDataType.xlGeneralFormat];

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyImport_ConnectionSettings_PreserveExtractionDefinition(bool web)
    {
        var batch = _fixture.BatchToken;
        string sheetName = _fixture.CreateTestSheet(batch);
        string queryName = UniqueConnectionName("LegacySettings");
        string connectionName = CreateLegacyImport(web, sheetName, queryName);
        _fixture.RegisterConnectionForCleanup(connectionName);
        RequireSuccess(_queryTables.SetProperties(batch, sheetName, queryName,
            backgroundQuery: true, refreshOnFileOpen: true, refreshPeriod: 17));
        var before = RequireSuccess(_queryTables.View(batch, sheetName, queryName));
        var viewed = RequireSuccess(_connections.View(batch, connectionName));
        Assert.Equal(web ? "WEB" : "TEXT", viewed.Type);
        Assert.Equal(before.Connection, viewed.ConnectionString);
        Assert.Null(viewed.CommandText);
        Assert.Null(viewed.CommandType);
        var listed = Assert.Single(RequireSuccess(_connections.List(batch)).Connections,
            item => item.Name == connectionName);
        Assert.Equal(viewed.Type, listed.Type);
        Assert.True(listed.BackgroundQuery);
        Assert.True(listed.RefreshOnFileOpen);
        var settings = RequireSuccess(_connections.GetProperties(batch, connectionName));
        Assert.True(settings.BackgroundQuery);
        Assert.True(settings.RefreshOnFileOpen);
        Assert.False(settings.SavePassword);
        Assert.Equal(17, settings.RefreshPeriod);

        var error = Assert.Throws<ArgumentException>(() => _connections.SetProperties(batch, connectionName,
            commandText: "SELECT * FROM SomethingElse", description: "Must not be applied",
            refreshPeriod: 99));
        Assert.Contains("SQL commandText", error.Message, StringComparison.Ordinal);
        Assert.Equal(viewed.DefinitionJson, RequireSuccess(_connections.View(batch, connectionName)).DefinitionJson);

        RequireSuccess(_connections.SetProperties(batch, connectionName,
            backgroundQuery: false, refreshOnFileOpen: false, savePassword: false, refreshPeriod: 31));
        settings = RequireSuccess(_connections.GetProperties(batch, connectionName));
        Assert.False(settings.BackgroundQuery);
        Assert.False(settings.RefreshOnFileOpen);
        Assert.False(settings.SavePassword);
        Assert.Equal(31, settings.RefreshPeriod);
        var after = RequireSuccess(_queryTables.View(batch, sheetName, queryName));
        Assert.Equal(before.Connection, after.Connection);
        Assert.Equal(before.Delimiter, after.Delimiter);
        Assert.Equal(before.Encoding, after.Encoding);
        Assert.Equal(before.WebSelectionType, after.WebSelectionType);
        Assert.Equal(before.WebTables, after.WebTables);
        Assert.Equal(before.WebFormatting, after.WebFormatting);
        Assert.False(after.BackgroundQuery);
        Assert.False(after.RefreshOnFileOpen);
        Assert.Equal(31, after.RefreshPeriod);
        RequireSuccess(_queryTables.Refresh(batch, sheetName, queryName));
        var rows = RequireSuccess(_commands.GetValues(batch, sheetName, "A1:B2")).Values;
        Assert.Equal(["Name", "Amount"], rows[0]);
        Assert.Equal("Selected", rows[1][0]);
        Assert.Equal(17d, Convert.ToDouble(rows[1][1], CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void LegacyImport_LoadTo_PreservesExtractionDefinition(bool web, bool fixedWidth)
    {
        var batch = _fixture.BatchToken;
        string sourceSheet = _fixture.CreateTestSheet(batch);
        string targetSheet = _fixture.CreateTestSheet(batch);
        string queryName = UniqueConnectionName("LegacyLoad");
        string connectionName = CreateLegacyImport(web, sourceSheet, queryName, fixedWidth);
        _fixture.RegisterConnectionForCleanup(connectionName);
        RequireSuccess(_queryTables.SetProperties(batch, sourceSheet, queryName,
            backgroundQuery: false, refreshOnFileOpen: true, refreshPeriod: 17,
            adjustColumnWidth: false, preserveFormatting: false));
        var before = RequireSuccess(_queryTables.View(batch, sourceSheet, queryName));
        string extractionBefore = ReadLegacyExtraction(sourceSheet, queryName, web);
        RequireSuccess(_connections.SetProperties(batch, connectionName, description: "Preserved import"));
        RequireSuccess(_commands.SetValues(batch, targetSheet, "G7", [["Unrelated"]]));
        RequireSuccess(_commands.SetValues(batch, sourceSheet, "G7", [["Source retained"]]));

        var result = RequireSuccess(_connections.LoadTo(batch, connectionName, targetSheet));
        Assert.Equal(_fixture.WorkbookPath, result.FilePath);

        var queries = RequireSuccess(_queryTables.List(batch)).QueryTables;
        var loaded = Assert.Single(queries, query => query.SheetName == targetSheet);
        Assert.Equal(connectionName.Replace(" ", "_"), loaded.Name);
        Assert.DoesNotContain(queries, query => query.SheetName == sourceSheet && query.Name == queryName);
        var after = RequireSuccess(_queryTables.View(batch, targetSheet, loaded.Name));
        Assert.Equal(extractionBefore, ReadLegacyExtraction(targetSheet, loaded.Name, web));
        Assert.Equal("Preserved import",
            Assert.Single(RequireSuccess(_connections.List(batch)).Connections,
                connection => connection.Name == connectionName).Description);
        Assert.Equal(before.Destination, after.Destination);
        Assert.Equal(before.SourceType, after.SourceType);
        Assert.Equal(before.Connection, after.Connection);
        Assert.Equal(before.Delimiter, after.Delimiter);
        Assert.Equal(before.Encoding, after.Encoding);
        Assert.Equal(before.WebSelectionType, after.WebSelectionType);
        Assert.Equal(before.WebTables, after.WebTables);
        Assert.Equal(before.WebFormatting, after.WebFormatting);
        Assert.Equal(before.BackgroundQuery, after.BackgroundQuery);
        Assert.Equal(before.RefreshOnFileOpen, after.RefreshOnFileOpen);
        Assert.Equal(before.RefreshPeriod, after.RefreshPeriod);
        Assert.Equal(before.AdjustColumnWidth, after.AdjustColumnWidth);
        Assert.Equal(before.PreserveFormatting, after.PreserveFormatting);
        RequireSuccess(_connections.Refresh(batch, connectionName));
        var rows = RequireSuccess(_commands.GetValues(batch, targetSheet, "A1:B2")).Values;
        Assert.Equal(["Name", "Amount"], rows[0]);
        Assert.Equal("Selected", rows[1][0]);
        Assert.Equal(17d, Convert.ToDouble(rows[1][1], CultureInfo.InvariantCulture));
        Assert.Equal("Unrelated",
            RequireSuccess(_commands.GetValues(batch, targetSheet, "G7")).Values[0][0]);
        Assert.Equal("Source retained",
            RequireSuccess(_commands.GetValues(batch, sourceSheet, "G7")).Values[0][0]);
        var original = RequireSuccess(_commands.GetValues(batch, sourceSheet, "A1:B2")).Values;
        Assert.Equal(["Name", "Amount"], original[0]);
        Assert.Equal("Selected", original[1][0]);
        Assert.Equal(17d, Convert.ToDouble(original[1][1], CultureInfo.InvariantCulture));
        string sourcePath = web ? new Uri(before.Connection[4..]).LocalPath : before.Connection[5..];
        File.WriteAllText(sourcePath, File.ReadAllText(sourcePath)
            .Replace("Selected", "Reloaded", StringComparison.Ordinal)
            .Replace("17", "29", StringComparison.Ordinal));
        RequireSuccess(_connections.Refresh(batch, connectionName));
        var refreshed = RequireSuccess(_commands.GetValues(batch, targetSheet, "A1:B2")).Values;
        Assert.Equal(["Name", "Amount"], refreshed[0]);
        Assert.Equal("Reloaded", refreshed[1][0]);
        Assert.Equal(29d, Convert.ToDouble(refreshed[1][1], CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyImport_LoadTo_ProtectedSourcePreservesItsCells(bool web)
    {
        var batch = _fixture.BatchToken;
        string sourceSheet = _fixture.CreateTestSheet(batch);
        string targetSheet = _fixture.CreateTestSheet(batch);
        string queryName = UniqueConnectionName("ProtectedSource");
        string connectionName = CreateLegacyImport(web, sourceSheet, queryName);
        _fixture.RegisterConnectionForCleanup(connectionName);
        RequireSuccess(_commands.SetValues(batch, targetSheet, "A1", [["Retained target"]]));
        var sheets = _fixture.CreateCommands<ISheetStyleCommands>();
        RequireSuccess(sheets.SetProtection(batch, sourceSheet, true));
        try
        {
            RequireSuccess(_connections.LoadTo(batch, connectionName, targetSheet));
            Assert.Single(RequireSuccess(_queryTables.List(batch)).QueryTables,
                query => query.SheetName == targetSheet);
            Assert.DoesNotContain(RequireSuccess(_queryTables.List(batch)).QueryTables,
                query => query.SheetName == sourceSheet);
            Assert.Equal("Selected",
                RequireSuccess(_commands.GetValues(batch, targetSheet, "A2")).Values[0][0]);
            Assert.Equal("Selected",
                RequireSuccess(_commands.GetValues(batch, sourceSheet, "A2")).Values[0][0]);
        }
        finally
        {
            RequireSuccess(sheets.SetProtection(batch, sourceSheet, false));
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LegacyImport_LoadTo_SameSheetRefreshesAtA1(bool web, bool move)
    {
        var batch = _fixture.BatchToken;
        string sheet = _fixture.CreateTestSheet(batch);
        string queryName = UniqueConnectionName("SameSheet");
        string connectionName = CreateLegacyImport(web, sheet, queryName, destination: move ? "D5" : "A1");
        _fixture.RegisterConnectionForCleanup(connectionName);
        var before = RequireSuccess(_queryTables.View(batch, sheet, queryName));
        string extraction = ReadLegacyExtraction(sheet, queryName, web);
        RequireSuccess(_commands.SetValues(batch, sheet, "G7", [["Unrelated"]]));
        string sourcePath = web ? new Uri(before.Connection[4..]).LocalPath : before.Connection[5..];
        File.WriteAllText(sourcePath, File.ReadAllText(sourcePath)
            .Replace("Selected", "Reloaded", StringComparison.Ordinal)
            .Replace("17", "29", StringComparison.Ordinal));

        RequireSuccess(_connections.LoadTo(batch, connectionName, sheet));

        var loaded = Assert.Single(RequireSuccess(_queryTables.List(batch)).QueryTables,
            query => query.SheetName == sheet);
        Assert.Equal(extraction, ReadLegacyExtraction(sheet, loaded.Name, web));
        Assert.Equal("Reloaded",
            RequireSuccess(_commands.GetValues(batch, sheet, "A2")).Values[0][0]);
        Assert.Equal(29d, Convert.ToDouble(
            RequireSuccess(_commands.GetValues(batch, sheet, "B2")).Values[0][0], CultureInfo.InvariantCulture));
        Assert.Equal("Unrelated",
            RequireSuccess(_commands.GetValues(batch, sheet, "G7")).Values[0][0]);
        Assert.Single(RequireSuccess(_connections.List(batch)).Connections,
            connection => connection.Name == connectionName);
        if (move)
            Assert.Equal("Selected",
                RequireSuccess(_commands.GetValues(batch, sheet, "D6")).Values[0][0]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyImport_LoadTo_ProtectedTargetPreservesSourceAndDestination(bool web)
    {
        var batch = _fixture.BatchToken;
        string sourceSheet = _fixture.CreateTestSheet(batch);
        string targetSheet = _fixture.CreateTestSheet(batch);
        string queryName = UniqueConnectionName("ProtectedLegacy");
        string connectionName = CreateLegacyImport(web, sourceSheet, queryName);
        _fixture.RegisterConnectionForCleanup(connectionName);
        var before = RequireSuccess(_queryTables.View(batch, sourceSheet, queryName));
        string extraction = ReadLegacyExtraction(sourceSheet, queryName, web);
        RequireSuccess(_commands.SetValues(batch, targetSheet, "A1", [["Retained target"]]));
        var sheets = _fixture.CreateCommands<ISheetStyleCommands>();
        RequireSuccess(sheets.SetProtection(batch, targetSheet, true));
        try
        {
            var error = Assert.Throws<InvalidOperationException>(() =>
                _connections.LoadTo(batch, connectionName, targetSheet));
            Assert.Contains("protected", error.Message, StringComparison.Ordinal);
            Assert.Contains("No QueryTables have been changed", error.Message, StringComparison.Ordinal);
            Assert.Equal(JsonSerializer.Serialize(before),
                JsonSerializer.Serialize(RequireSuccess(_queryTables.View(batch, sourceSheet, queryName))));
            Assert.Equal(extraction, ReadLegacyExtraction(sourceSheet, queryName, web));
            Assert.DoesNotContain(RequireSuccess(_queryTables.List(batch)).QueryTables,
                query => query.SheetName == targetSheet);
            Assert.Equal("Retained target",
                RequireSuccess(_commands.GetValues(batch, targetSheet, "A1")).Values[0][0]);
            RequireSuccess(_connections.Refresh(batch, connectionName));
            Assert.Equal("Selected",
                RequireSuccess(_commands.GetValues(batch, sourceSheet, "A2")).Values[0][0]);
        }
        finally
        {
            RequireSuccess(sheets.SetProtection(batch, targetSheet, false));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyImport_LoadTo_RefreshFailureReportsSurvivingImports(bool web)
    {
        var batch = _fixture.BatchToken;
        string sourceSheet = _fixture.CreateTestSheet(batch);
        string targetSheet = _fixture.CreateTestSheet(batch);
        string queryName = UniqueConnectionName("MissingLegacy");
        string connectionName = CreateLegacyImport(web, sourceSheet, queryName);
        _fixture.RegisterConnectionForCleanup(connectionName);
        var before = RequireSuccess(_queryTables.View(batch, sourceSheet, queryName));
        string extraction = ReadLegacyExtraction(sourceSheet, queryName, web);
        var connectionNames = RequireSuccess(_connections.List(batch)).Connections
            .Select(connection => connection.Name).ToHashSet(StringComparer.Ordinal);
        RequireSuccess(_commands.SetValues(batch, targetSheet, "A1", [["Retained target"]]));
        RequireSuccess(_commands.SetValues(batch, targetSheet, "G7", [["Unrelated"]]));
        string sourcePath = web ? new Uri(before.Connection[4..]).LocalPath : before.Connection[5..];
        File.Delete(sourcePath);

        var error = Assert.Throws<InvalidOperationException>(() =>
            _connections.LoadTo(batch, connectionName, targetSheet));

        var copiedConnection = Assert.Single(RequireSuccess(_connections.List(batch)).Connections,
            connection => !connectionNames.Contains(connection.Name));
        _fixture.RegisterConnectionForCleanup(copiedConnection.Name);
        var queries = RequireSuccess(_queryTables.List(batch)).QueryTables;
        var original = Assert.Single(queries, query => query.SheetName == sourceSheet);
        var copied = Assert.Single(queries, query => query.SheetName == targetSheet);
        Assert.Equal(queryName, original.Name);
        Assert.Equal(extraction, ReadLegacyExtraction(sourceSheet, original.Name, web));
        Assert.Equal(extraction, ReadLegacyExtraction(targetSheet, copied.Name, web));
        Assert.Single(RequireSuccess(_connections.List(batch)).Connections,
            connection => connection.Name == connectionName);
        Assert.Equal("Selected",
            RequireSuccess(_commands.GetValues(batch, sourceSheet, "A2")).Values[0][0]);
        Assert.Equal("Retained target",
            RequireSuccess(_commands.GetValues(batch, targetSheet, "A1")).Values[0][0]);
        Assert.Equal("Unrelated",
            RequireSuccess(_commands.GetValues(batch, targetSheet, "G7")).Values[0][0]);
        Assert.Contains(connectionName, error.Message, StringComparison.Ordinal);
        Assert.Contains(queryName, error.Message, StringComparison.Ordinal);
        Assert.Contains(sourceSheet, error.Message, StringComparison.Ordinal);
        Assert.Contains(copied.Name, error.Message, StringComparison.Ordinal);
        Assert.Contains(copiedConnection.Name, error.Message, StringComparison.Ordinal);
        Assert.Contains(targetSheet, error.Message, StringComparison.Ordinal);
        Assert.Contains("remains", error.Message, StringComparison.Ordinal);
    }

    private string ReadLegacyExtraction(string sheetName, string queryName, bool web) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.QueryTables? queries = null;
            Excel.QueryTable? query = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                queries = sheet.QueryTables;
                query = queries.Item(queryName);
                object extraction = web ? new
                {
                    query.WebSelectionType,
                    query.WebTables,
                    query.WebFormatting,
                    query.WebDisableDateRecognition,
                    query.WebDisableRedirections,
                    query.WebSingleBlockTextImport,
                    query.WebPreFormattedTextToColumns,
                    query.WebConsecutiveDelimitersAsOne
                } : (object)new
                {
                    query.TextFileParseType,
                    query.TextFileStartRow,
                    query.TextFilePlatform,
                    query.TextFileTextQualifier,
                    query.TextFileTabDelimiter,
                    query.TextFileSemicolonDelimiter,
                    query.TextFileCommaDelimiter,
                    query.TextFileSpaceDelimiter,
                    query.TextFileOtherDelimiter,
                    query.TextFileConsecutiveDelimiter,
                    query.TextFileDecimalSeparator,
                    query.TextFileThousandsSeparator,
                    query.TextFileTrailingMinusNumbers,
                    ColumnTypes = ((Array)(object)query.TextFileColumnDataTypes).Cast<object>().Select(Convert.ToInt32).ToArray(),
                    Widths = query.TextFileParseType == Excel.XlTextParsingType.xlFixedWidth
                        ? ((Array)(object)query.TextFileFixedColumnWidths).Cast<object>().Select(Convert.ToInt32).ToArray()
                        : null
                };
                return JsonSerializer.Serialize(new
                {
                    extraction,
                    query.RefreshStyle,
                    query.SaveData,
                    query.SavePassword,
                    query.PreserveColumnInfo,
                    query.FieldNames,
                    query.RowNumbers,
                    query.FillAdjacentFormulas
                });
            }
            finally
            {
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private string CreateLegacyImport(
        bool web, string sheetName, string queryName, bool fixedWidth = false, string destination = "A1")
    {
        var batch = _fixture.BatchToken;
        string source = _fixture.CreateInputFile(web ? ".html" : ".csv", web
            ? "<html><body><table><tr><td>Ignored</td></tr></table>" +
              "<table><tr><th>Name</th><th>Amount</th></tr><tr><td>Selected</td><td>17</td></tr></table></body></html>"
            : fixedWidth ? "Name    Amount\r\nSelected17\r\n" : "Name;Amount\r\nSelected;17\r\n");
        if (web)
            RequireSuccess(_queryTables.CreateWeb(batch, queryName, new Uri(source).AbsoluteUri,
                sheetName, destination, selectionType: "specified-tables", webTables: "2"));
        else
            RequireSuccess(_queryTables.CreateText(batch, queryName, source, sheetName, destination, delimiter: ";"));

        return _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.QueryTables? queries = null;
            Excel.QueryTable? query = null;
            Excel.WorkbookConnection? connection = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                queries = sheet.QueryTables;
                query = queries.Item(queryName);
                if (fixedWidth)
                {
                    query.TextFileParseType = Excel.XlTextParsingType.xlFixedWidth;
                    query.TextFileFixedColumnWidths = s_legacyFixedWidths;
                    query.TextFileColumnDataTypes = s_legacyFixedTypes;
                    Assert.True(query.Refresh(false));
                }
                connection = query.WorkbookConnection;
                return connection.Name;
            }
            finally
            {
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queries);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }
}
