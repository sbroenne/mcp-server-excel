using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "QueryTable")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceQueryTableTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IQueryTableCommands _queryTables =
        ServiceCommandProxy.Create<IQueryTableCommands>(fixture);

    [Fact]
    public void TextImport_LifecycleAndConfiguration_RoundTrips()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var sourcePath = _fixture.CreateInputFile(
            ".csv", "Name,Value\nCafé,10\nBeta,20\n");

        var createResult = _queryTables.CreateText(
            batch, "CsvImport", sourcePath, sheetName, "B2",
            delimiter: ",", textQualifier: "double-quote", encoding: 65001,
            hasHeaders: true);
        Assert.True(createResult.Success);

        var listed = Assert.Single(_queryTables.List(batch).QueryTables);
        Assert.Equal("CsvImport", listed.Name);
        Assert.Equal(sheetName, listed.SheetName);
        Assert.Equal("B2", listed.Destination);
        Assert.Equal("text", listed.SourceType);

        var viewResult = _queryTables.View(batch, sheetName, "CsvImport");
        Assert.True(viewResult.Success);
        Assert.Equal(",", viewResult.Delimiter);
        Assert.NotNull(viewResult.Encoding);
        Assert.Equal(
            "Café",
            _commands.GetValues(batch, sheetName, "B3").Values[0][0]);

        Assert.True(_queryTables.SetProperties(
            batch, sheetName, "CsvImport", backgroundQuery: false,
            refreshOnFileOpen: true, refreshPeriod: 15,
            adjustColumnWidth: false, preserveFormatting: true).Success);

        var configured = _queryTables.View(batch, sheetName, "CsvImport");
        Assert.False(configured.BackgroundQuery);
        Assert.True(configured.RefreshOnFileOpen);
        Assert.Equal(15, configured.RefreshPeriod);
        Assert.False(configured.AdjustColumnWidth);
        Assert.True(configured.PreserveFormatting);

        var status = _queryTables.GetRefreshStatus(batch, sheetName, "CsvImport");
        Assert.True(status.Success);
        Assert.False(status.IsRefreshing);

        var cancelResult = _queryTables.CancelRefresh(batch, sheetName, "CsvImport");
        Assert.True(cancelResult.Success);
        Assert.False(cancelResult.WasRefreshing);
        Assert.True(_queryTables.Refresh(batch, sheetName, "CsvImport").Success);
        Assert.True(_queryTables.Delete(batch, sheetName, "CsvImport").Success);
        Assert.Empty(_queryTables.List(batch).QueryTables);
    }

    [Fact]
    public void WebImport_FromLocalHtml_RoundTrips()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var htmlPath = _fixture.CreateInputFile(
            ".html",
            "<html><body><table><tr><th>Name</th><th>Value</th></tr>" +
            "<tr><td>Alpha</td><td>10</td></tr></table></body></html>");

        var createResult = _queryTables.CreateWeb(
            batch, "HtmlImport", new Uri(htmlPath).AbsoluteUri,
            sheetName, "A1", selectionType: "specified-tables",
            webTables: "1", formatting: "none");
        Assert.True(createResult.Success);

        var viewResult = _queryTables.View(batch, sheetName, "HtmlImport");
        Assert.True(viewResult.Success);
        Assert.Equal("web", viewResult.SourceType);
        Assert.Equal("specified-tables", viewResult.WebSelectionType);
        Assert.Equal("1", viewResult.WebTables);
        Assert.Equal("none", viewResult.WebFormatting);

        Assert.True(_queryTables.Delete(batch, sheetName, "HtmlImport").Success);
        Assert.Empty(_queryTables.List(batch).QueryTables);
    }

    [Theory]
    [InlineData("ftp://example.com/data.html")]
    [InlineData("mailto:user@example.com")]
    [InlineData("custom-fetch://example.com/data")]
    public void WebImport_UnsupportedScheme_ThrowsBeforeCom(string url)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var exception = Assert.Throws<ArgumentException>(() =>
            _queryTables.CreateWeb(
                batch, "UnsupportedImport", url, sheetName, "A1"));

        Assert.Contains(
            "HTTP, HTTPS, or file URI", exception.Message,
            StringComparison.Ordinal);
        Assert.Empty(_queryTables.List(batch).QueryTables);
    }
}
