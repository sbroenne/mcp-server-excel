using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration.Session;

public partial class ExcelBatchTests
{
    private static readonly int[] s_textExtractionColumnTypes =
    [
        (int)Excel.XlColumnDataType.xlTextFormat,
        (int)Excel.XlColumnDataType.xlSkipColumn,
        (int)Excel.XlColumnDataType.xlGeneralFormat
    ];
    private static readonly int[] s_changedTextExtractionColumnTypes =
    [
        (int)Excel.XlColumnDataType.xlTextFormat,
        (int)Excel.XlColumnDataType.xlTextFormat,
        (int)Excel.XlColumnDataType.xlSkipColumn
    ];

    [Fact]
    public void Execute_UnrefreshedOleDbDefinitions_ExposeExactUnsetRefreshDateFailure()
    {
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Excel.Connections? connections = null;
            Excel.WorkbookConnection? connection = null;
            Excel.OLEDBConnection? provider = null;
            try
            {
                connections = ctx.Book.Connections;
                connection = connections.Add2("TypedUnsetRefreshDate", "Unset timestamp probe",
                    "OLEDB;Provider=MSOLAP.8;Data Source=UnconfiguredRegressionSource;Initial Catalog=RegressionCube",
                    "Model", Excel.XlCmdType.xlCmdCube, false, false);
                provider = connection.OLEDBConnection;
                Assert.True(provider.OLAP);
                Assert.False(provider.IsConnected);
                var error = Assert.Throws<COMException>(() => _ = provider.RefreshDate);
                Assert.Equal(unchecked((int)0x800A03EC), error.HResult);
                Assert.Equal("Model", (object)provider.CommandText);
                Assert.Equal(Excel.XlCmdType.xlCmdCube, provider.CommandType);
                Assert.False(provider.Refreshing);
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref provider);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref connections);
            }
        });
        AssertNoExternalInteropAssemblies();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Execute_LegacyConnectionDefinitions_ExposeQueryTableSettings(bool web)
    {
        string source = Path.Combine(Path.GetTempPath(), $"typed-connection-{Guid.NewGuid():N}.csv");
        _temporaryFiles.Add(source);
        File.WriteAllText(source, "Region,Amount\r\nNorth,17\r\nSouth,29\r\n");
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.QueryTables? queries = null;
            Excel.Range? destination = null;
            Excel.QueryTable? query = null;
            Excel.WorkbookConnection? connection = null;
            Excel.TextConnection? text = null;
            Excel.Ranges? ranges = null;
            Excel.Range? connectedRange = null;
            Excel.QueryTable? linked = null;
            Excel.WorkbookConnection? linkedConnection = null;
            Excel.Workbook? parent = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                destination = sheet.Range["A1"];
                queries = sheet.QueryTables;
                query = queries.Add(web ? "URL;http://127.0.0.1/typed-probe" : $"TEXT;{source}", destination);
                query.Name = web ? "TypedWebProbe" : "TypedTextProbe";
                query.BackgroundQuery = false;
                query.RefreshOnFileOpen = true;
                query.RefreshPeriod = 17;
                query.SavePassword = false;
                if (!web)
                {
                    query.TextFileParseType = Excel.XlTextParsingType.xlDelimited;
                    query.TextFileCommaDelimiter = true;
                    Assert.True(query.Refresh(false));
                }
                connection = query.WorkbookConnection;
                Assert.Equal(web ? Excel.XlConnectionType.xlConnectionTypeWEB : Excel.XlConnectionType.xlConnectionTypeTEXT,
                    connection.Type);
                Assert.Equal(web ? 5 : 4, (int)connection.Type);
                Assert.False(query.BackgroundQuery);
                Assert.True(query.RefreshOnFileOpen);
                Assert.Equal(17, query.RefreshPeriod);
                Assert.False(query.SavePassword);
                parent = (Excel.Workbook)connection.Parent;
                Assert.Equal(ctx.Book.Name, parent.Name);
                var commandError = Assert.Throws<OutOfMemoryException>(() => _ = (object?)query.CommandText);
                Assert.Equal(unchecked((int)0x8007000E), commandError.HResult);
                string[] connectionMembers = ReadNativeConnectionMembers(connection);
                Assert.Contains("Type", connectionMembers);
                Assert.Contains("TextConnection", connectionMembers);
                Assert.DoesNotContain("WebConnection", connectionMembers);
                Assert.DoesNotContain("ConnectionString", connectionMembers);
                if (!web)
                {
                    text = connection.TextConnection;
                    object value = text.Connection;
                    Assert.Contains(Path.GetFileName(source), value.ToString(), StringComparison.OrdinalIgnoreCase);
                    string[] members = ReadNativeConnectionMembers(text);
                    Assert.Contains("Connection", members);
                    Assert.DoesNotContain("BackgroundQuery", members);
                    Assert.DoesNotContain("RefreshOnFileOpen", members);
                    Assert.DoesNotContain("RefreshPeriod", members);
                    Assert.DoesNotContain("SavePassword", members);
                    Assert.DoesNotContain("CommandText", members);
                    ranges = connection.Ranges;
                    Assert.Equal(1, ranges.Count);
                    connectedRange = ranges[1];
                    linked = connectedRange.QueryTable;
                    linkedConnection = linked.WorkbookConnection;
                    Assert.Equal(connection.Name, linkedConnection.Name);
                    Assert.Equal(query.Name, linked.Name);
                    Assert.Equal(query.BackgroundQuery, linked.BackgroundQuery);
                    Assert.Equal(query.RefreshOnFileOpen, linked.RefreshOnFileOpen);
                    Assert.Equal(query.RefreshPeriod, linked.RefreshPeriod);
                    Assert.Equal(query.SavePassword, linked.SavePassword);
                    Assert.Equal("Region", (object)destination.Value2);
                }
                _output.WriteLine($"{(web ? "WEB" : "TEXT")} native type {(int)connection.Type}; " +
                    "refresh settings are exposed on QueryTable, not TextConnection or WebConnection.");
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref linkedConnection);
                ComUtilities.Release(ref parent);
                ComUtilities.Release(ref linked);
                ComUtilities.Release(ref connectedRange);
                ComUtilities.Release(ref ranges);
                ComUtilities.Release(ref text);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref queries);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        AssertNoExternalInteropAssemblies();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Execute_LegacyExtractionDefinitions_PreserveSelectionAndRefresh(bool web)
    {
        string source = Path.Combine(Path.GetTempPath(),
            $"typed-extraction-{Guid.NewGuid():N}.{(web ? "html" : "csv")}");
        _temporaryFiles.Add(source);
        string originalSource = web
            ? "<html><body><table><tr><th>Ignored</th></tr><tr><td>not selected</td></tr></table>" +
              "<table><tr><th>Id</th><th>Amount</th></tr>" +
              "<tr><td>Selected</td><td>17</td></tr></table></body></html>"
            : "not imported\r\nId;Ignored;Amount\r\n0007;not selected;17\r\n0008;not selected;29\r\n";
        File.WriteAllText(source, originalSource);
        using var batch = ExcelSession.BeginBatch(_testFileCopy!);
        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? originalSheet = null;
            Excel.Worksheet? sheet = null;
            Excel.Worksheet? activeSheet = null;
            Excel.QueryTables? queries = null;
            Excel.Range? destination = null;
            Excel.Range? result = null;
            Excel.QueryTable? query = null;
            Excel.WorkbookConnection? connection = null;
            Excel.Ranges? ranges = null;
            Excel.Range? connectedRange = null;
            Excel.QueryTable? linked = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                originalSheet = (Excel.Worksheet)sheets[1];
                sheet = (Excel.Worksheet)sheets.Add();
                sheet.Name = "ExtractionDefinition";
                queries = sheet.QueryTables;
                destination = sheet.Range["B2"];
                string expectedConnection = web ? $"URL;{new Uri(source).AbsoluteUri}" : $"TEXT;{source}";
                query = queries.Add(expectedConnection, destination);
                query.Name = "TypedExtractionProbe";
                query.BackgroundQuery = false;
                if (web)
                {
                    query.WebSelectionType = Excel.XlWebSelectionType.xlSpecifiedTables;
                    query.WebTables = "2";
                    query.WebFormatting = Excel.XlWebFormatting.xlWebFormattingNone;
                }
                else
                {
                    query.TextFileParseType = Excel.XlTextParsingType.xlDelimited;
                    query.TextFileStartRow = 2;
                    query.TextFileSemicolonDelimiter = true;
                    query.TextFileCommaDelimiter = false;
                    query.TextFileTabDelimiter = false;
                    query.TextFileTextQualifier = Excel.XlTextQualifier.xlTextQualifierDoubleQuote;
                    query.TextFileColumnDataTypes = s_textExtractionColumnTypes;
                }

                originalSheet.Activate();
                activeSheet = (Excel.Worksheet)ctx.Book.ActiveSheet;
                Assert.Equal(originalSheet.Name, activeSheet.Name);
                Assert.NotEqual(sheet.Name, activeSheet.Name);
                ComUtilities.Release(ref activeSheet);
                var inactiveError = Assert.Throws<OutOfMemoryException>(() => _ = (object?)query.CommandText);
                Assert.Equal(unchecked((int)0x8007000E), inactiveError.HResult);

                sheet.Activate();
                activeSheet = (Excel.Worksheet)ctx.Book.ActiveSheet;
                Assert.Equal(sheet.Name, activeSheet.Name);
                ComUtilities.Release(ref activeSheet);
                var activeError = Assert.Throws<OutOfMemoryException>(() => _ = (object?)query.CommandText);
                Assert.Equal(unchecked((int)0x8007000E), activeError.HResult);
                Assert.Equal(expectedConnection, (object)query.Connection);
                Assert.Equal(web ? Excel.XlQueryType.xlWebQuery : Excel.XlQueryType.xlTextImport, query.QueryType);
                Assert.True(query.Refresh(false));

                if (web)
                {
                    Assert.Equal(Excel.XlWebSelectionType.xlSpecifiedTables, query.WebSelectionType);
                    Assert.Equal("2", query.WebTables);
                    Assert.Equal(Excel.XlWebFormatting.xlWebFormattingNone, query.WebFormatting);
                }
                else
                {
                    Assert.Equal(Excel.XlTextParsingType.xlDelimited, query.TextFileParseType);
                    Assert.Equal(2, query.TextFileStartRow);
                    Assert.True(query.TextFileSemicolonDelimiter);
                    Assert.False(query.TextFileCommaDelimiter);
                    Assert.False(query.TextFileTabDelimiter);
                    Assert.Equal(Excel.XlTextQualifier.xlTextQualifierDoubleQuote, query.TextFileTextQualifier);
                    var types = Assert.IsAssignableFrom<Array>((object)query.TextFileColumnDataTypes);
                    Assert.Equal(s_textExtractionColumnTypes, types.Cast<object>().Select(Convert.ToInt32));
                }

                result = query.ResultRange;
                var values = Assert.IsType<object[,]>((object)result.Value2);
                Assert.Equal(2, values.GetLength(1));
                Assert.Equal(web ? 2 : 3, values.GetLength(0));
                Assert.Equal("Id", values[1, 1]);
                Assert.Equal("Amount", values[1, 2]);
                Assert.Equal(web ? "Selected" : "0007", values[2, 1]);
                Assert.Equal(17d, values[2, 2]);
                if (!web)
                {
                    Assert.Equal("0008", values[3, 1]);
                    Assert.Equal(29d, values[3, 2]);
                }
                ComUtilities.Release(ref result);

                connection = query.WorkbookConnection;
                ranges = connection.Ranges;
                Assert.Equal(1, ranges.Count);
                connectedRange = ranges[1];
                linked = connectedRange.QueryTable;
                Assert.Equal(query.Name, linked.Name);
                Assert.Equal(expectedConnection, (object)linked.Connection);
                var refreshedError = Assert.Throws<OutOfMemoryException>(() => _ = (object?)linked.CommandText);
                Assert.Equal(unchecked((int)0x8007000E), refreshedError.HResult);
                Assert.Equal(originalSource, File.ReadAllText(source));

                if (web)
                {
                    linked.WebTables = "1";
                }
                else
                {
                    linked.TextFileColumnDataTypes = s_changedTextExtractionColumnTypes;
                }
                Assert.True(linked.Refresh(false));
                result = linked.ResultRange;
                values = Assert.IsType<object[,]>((object)result.Value2);
                Assert.Equal(web ? 1 : 2, values.GetLength(1));
                Assert.Equal(web ? "Ignored" : "Id", values[1, 1]);
                Assert.Equal(web ? "not selected" : "0007", values[2, 1]);
                if (web)
                {
                    Assert.Equal("1", query.WebTables);
                    Assert.Equal(2, values.GetLength(0));
                }
                else
                {
                    var changedTypes = Assert.IsAssignableFrom<Array>((object)query.TextFileColumnDataTypes);
                    Assert.Equal(s_changedTextExtractionColumnTypes, changedTypes.Cast<object>().Select(Convert.ToInt32));
                    Assert.Equal("Ignored", values[1, 2]);
                    Assert.Equal("not selected", values[2, 2]);
                    Assert.Equal("0008", values[3, 1]);
                    Assert.Equal(3, values.GetLength(0));
                }
                Assert.Equal(originalSource, File.ReadAllText(source));
                _output.WriteLine($"{(web ? "WEB" : "TEXT")}: active/inactive CommandText reads fail with 0x8007000E, " +
                    "but typed extraction settings read back, change, and control the refreshed cells.");
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref linked);
                ComUtilities.Release(ref connectedRange);
                ComUtilities.Release(ref ranges);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref result);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref queries);
                ComUtilities.Release(ref activeSheet);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref originalSheet);
                ComUtilities.Release(ref sheets);
            }
        });
        AssertNoExternalInteropAssemblies();
    }

    private static string[] ReadNativeConnectionMembers(object value)
    {
        ITypeInfo? information = null;
        try
        {
            ((IAutomationTypeInformation)value).GetTypeInfo(0, 0, out information);
            information.GetTypeAttr(out nint attributes);
            try
            {
                var type = Marshal.PtrToStructure<TYPEATTR>(attributes);
                var names = new HashSet<string>(StringComparer.Ordinal);
                for (int index = 0; index < type.cFuncs; index++)
                {
                    information.GetFuncDesc(index, out nint descriptor);
                    try
                    {
                        var function = Marshal.PtrToStructure<FUNCDESC>(descriptor);
                        information.GetDocumentation(function.memid, out string name, out _, out _, out _);
                        names.Add(name);
                    }
                    finally
                    {
                        information.ReleaseFuncDesc(descriptor);
                    }
                }
                return [.. names.Order(StringComparer.Ordinal)];
            }
            finally
            {
                information.ReleaseTypeAttr(attributes);
            }
        }
        finally
        {
            ComUtilities.Release(ref information);
        }
    }
}
