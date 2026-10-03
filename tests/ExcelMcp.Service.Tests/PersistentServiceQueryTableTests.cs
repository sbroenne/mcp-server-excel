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
        RequireSuccess(createResult);

        var listed = Assert.Single(RequireSuccess(_queryTables.List(batch)).QueryTables);
        Assert.Equal("CsvImport", listed.Name);
        Assert.Equal(sheetName, listed.SheetName);
        Assert.Equal("B2", listed.Destination);
        Assert.Equal("text", listed.SourceType);

        var viewResult = _queryTables.View(batch, sheetName, "CsvImport");
        RequireSuccess(viewResult);
        Assert.Equal(",", viewResult.Delimiter);
        // Native Excel reads back UTF-8 code page 65001 as platform 98.
        Assert.Equal(98, viewResult.Encoding);
        AssertImportedRows(sheetName, "B2:C4", "Café", 10, "Beta", 20);

        Assert.True(_queryTables.SetProperties(
            batch, sheetName, "CsvImport", backgroundQuery: false,
            refreshOnFileOpen: true, refreshPeriod: 15,
            adjustColumnWidth: false, preserveFormatting: true).Success);

        var configured = RequireSuccess(_queryTables.View(batch, sheetName, "CsvImport"));
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
        File.WriteAllText(sourcePath, "Name,Value\nGamma,30\nDelta,40\n");
        RequireSuccess(_queryTables.Refresh(batch, sheetName, "CsvImport"));
        AssertImportedRows(sheetName, "B2:C4", "Gamma", 30, "Delta", 40);
        Assert.False(RequireSuccess(_queryTables.GetRefreshStatus(batch, sheetName, "CsvImport")).IsRefreshing);
        RequireSuccess(_queryTables.Delete(batch, sheetName, "CsvImport"));
        Assert.Empty(RequireSuccess(_queryTables.List(batch)).QueryTables);
        AssertImportedRows(sheetName, "B2:C4", "Gamma", 30, "Delta", 40);
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
        RequireSuccess(createResult);

        var viewResult = _queryTables.View(batch, sheetName, "HtmlImport");
        RequireSuccess(viewResult);
        Assert.Equal("web", viewResult.SourceType);
        Assert.Equal("specified-tables", viewResult.WebSelectionType);
        Assert.Equal("1", viewResult.WebTables);
        Assert.Equal("none", viewResult.WebFormatting);

        var rows = RequireSuccess(_commands.GetValues(batch, sheetName, "A1:B2")).Values;
        Assert.Equal(["Name", "Value"], rows[0]);
        Assert.Equal("Alpha", rows[1][0]);
        Assert.Equal(10, Convert.ToDouble(rows[1][1], System.Globalization.CultureInfo.InvariantCulture));
        RequireSuccess(_queryTables.Delete(batch, sheetName, "HtmlImport"));
        Assert.Empty(RequireSuccess(_queryTables.List(batch)).QueryTables);
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(rows),
            System.Text.Json.JsonSerializer.Serialize(
                RequireSuccess(_commands.GetValues(batch, sheetName, "A1:B2")).Values));
    }

    [Theory]
    [InlineData("ftp://example.com/data.html")]
    [InlineData("mailto:user@example.com")]
    [InlineData("custom-fetch://example.com/data")]
    public void WebImport_UnsupportedScheme_ThrowsBeforeCom(string url)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:B1", [["retained", 7]]));
        var before = RequireSuccess(_queryTables.List(batch));

        var exception = Assert.Throws<ArgumentException>(() =>
            _queryTables.CreateWeb(
                batch, "UnsupportedImport", url, sheetName, "A1"));

        Assert.Contains(
            "HTTP, HTTPS, or file URI", exception.Message,
            StringComparison.Ordinal);
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(before.QueryTables),
            System.Text.Json.JsonSerializer.Serialize(RequireSuccess(_queryTables.List(batch)).QueryTables));
        Assert.Equal(["retained", 7],
            Assert.Single(RequireSuccess(_commands.GetValues(batch, sheetName, "A1:B1")).Values));
    }

    private void AssertImportedRows(string sheetName, string address,
        string firstName, double firstValue, string secondName, double secondValue)
    {
        var rows = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, address)).Values;
        Assert.Equal(3, rows.Count);
        Assert.All(rows, row => Assert.Equal(2, row.Count));
        Assert.Equal(["Name", "Value"], rows[0]);
        Assert.Equal(firstName, rows[1][0]);
        Assert.Equal(firstValue, Convert.ToDouble(rows[1][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(secondName, rows[2][0]);
        Assert.Equal(secondValue, Convert.ToDouble(rows[2][1], System.Globalization.CultureInfo.InvariantCulture));
    }
}
