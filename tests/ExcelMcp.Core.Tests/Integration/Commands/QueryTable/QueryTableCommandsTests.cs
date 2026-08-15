using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.QueryTable;

[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "QueryTable")]
[Trait("RequiresExcel", "true")]
public sealed class QueryTableCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly QueryTableCommands _commands = new();
    private readonly TempDirectoryFixture _fixture;

    public QueryTableCommandsTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void TextImport_LifecycleAndConfiguration_RoundTrips()
    {
        var workbookPath = _fixture.CreateTestFile();
        var sourcePath = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests),
            nameof(TextImport_LifecycleAndConfiguration_RoundTrips),
            _fixture.TempDir,
            ".csv",
            "Name,Value\nCafé,10\nBeta,20\n");

        using var batch = ExcelSession.BeginBatch(workbookPath);
        CreateSheet(batch, "Imports");

        var createResult = _commands.CreateText(
            batch,
            "CsvImport",
            sourcePath,
            "Imports",
            "B2",
            delimiter: ",",
            textQualifier: "double-quote",
            encoding: 65001,
            hasHeaders: true);
        Assert.True(createResult.Success);

        var listResult = _commands.List(batch);
        var listed = Assert.Single(listResult.QueryTables);
        Assert.Equal("CsvImport", listed.Name);
        Assert.Equal("Imports", listed.SheetName);
        Assert.Equal("B2", listed.Destination);
        Assert.Equal("text", listed.SourceType);

        var viewResult = _commands.View(batch, "Imports", "CsvImport");
        Assert.True(viewResult.Success);
        Assert.Equal(",", viewResult.Delimiter);
        Assert.NotNull(viewResult.Encoding);
        var importedValues = new RangeCommands().GetValues(batch, "Imports", "B3");
        Assert.Equal("Café", importedValues.Values[0][0]);

        var configureResult = _commands.SetProperties(
            batch,
            "Imports",
            "CsvImport",
            backgroundQuery: false,
            refreshOnFileOpen: true,
            refreshPeriod: 15,
            adjustColumnWidth: false,
            preserveFormatting: true);
        Assert.True(configureResult.Success);

        var configured = _commands.View(batch, "Imports", "CsvImport");
        Assert.False(configured.BackgroundQuery);
        Assert.True(configured.RefreshOnFileOpen);
        Assert.Equal(15, configured.RefreshPeriod);
        Assert.False(configured.AdjustColumnWidth);
        Assert.True(configured.PreserveFormatting);

        var status = _commands.GetRefreshStatus(batch, "Imports", "CsvImport");
        Assert.True(status.Success);
        Assert.False(status.IsRefreshing);

        var cancelResult = _commands.CancelRefresh(batch, "Imports", "CsvImport");
        Assert.True(cancelResult.Success);
        Assert.False(cancelResult.WasRefreshing);

        var refreshResult = _commands.Refresh(batch, "Imports", "CsvImport");
        Assert.True(refreshResult.Success);

        var deleteResult = _commands.Delete(batch, "Imports", "CsvImport");
        Assert.True(deleteResult.Success);
        Assert.Empty(_commands.List(batch).QueryTables);
    }

    [Fact]
    public void WebImport_FromLocalHtml_RoundTrips()
    {
        var workbookPath = _fixture.CreateTestFile();
        var htmlPath = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests),
            nameof(WebImport_FromLocalHtml_RoundTrips),
            _fixture.TempDir,
            ".html",
            "<html><body><table><tr><th>Name</th><th>Value</th></tr><tr><td>Alpha</td><td>10</td></tr></table></body></html>");

        using var batch = ExcelSession.BeginBatch(workbookPath);
        CreateSheet(batch, "WebImports");

        var createResult = _commands.CreateWeb(
            batch,
            "HtmlImport",
            new Uri(htmlPath).AbsoluteUri,
            "WebImports",
            "A1",
            selectionType: "specified-tables",
            webTables: "1",
            formatting: "none");
        Assert.True(createResult.Success);

        var viewResult = _commands.View(batch, "WebImports", "HtmlImport");
        Assert.True(viewResult.Success);
        Assert.Equal("web", viewResult.SourceType);
        Assert.Equal("specified-tables", viewResult.WebSelectionType);
        Assert.Equal("1", viewResult.WebTables);
        Assert.Equal("none", viewResult.WebFormatting);

        var deleteResult = _commands.Delete(batch, "WebImports", "HtmlImport");
        Assert.True(deleteResult.Success);
        Assert.Empty(_commands.List(batch).QueryTables);
    }

    [Theory]
    [InlineData("ftp://example.com/data.html")]
    [InlineData("mailto:user@example.com")]
    [InlineData("custom-fetch://example.com/data")]
    public void WebImport_UnsupportedScheme_ThrowsBeforeCom(string url)
    {
        var workbookPath = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(workbookPath);
        CreateSheet(batch, "WebImports");

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.CreateWeb(
                batch,
                "UnsupportedImport",
                url,
                "WebImports",
                "A1"));

        Assert.Contains("HTTP, HTTPS, or file URI", exception.Message, StringComparison.Ordinal);
        Assert.Empty(_commands.List(batch).QueryTables);
    }

    private static void CreateSheet(IExcelBatch batch, string sheetName)
    {
        batch.Execute((ctx, ct) =>
        {
            var sheet = ctx.Book.Worksheets.Add();
            sheet.Name = sheetName;
            return 0;
        });
    }
}
