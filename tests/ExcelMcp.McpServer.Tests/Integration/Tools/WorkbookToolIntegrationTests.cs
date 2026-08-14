using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end tests for generated workbook MCP operations using real Excel.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "true")]
public class WorkbookToolIntegrationTests(ITestOutputHelper output)
    : McpIntegrationTestBase(output, "WorkbookToolIntegrationClient")
{
    private string _tempDirectory = string.Empty;
    private string _workbookPath = string.Empty;
    private string _sessionId = string.Empty;

    protected override async Task InitializeTestAsync()
    {
        _tempDirectory = CreateTempDirectory("WorkbookToolIntegration");
        _workbookPath = Path.Join(_tempDirectory, "WorkbookTool.xlsx");
        _sessionId = await CreateWorkbookSessionAsync(_workbookPath);
    }

    [Fact]
    public async Task DocumentProperty_SetAndGet_UsesGeneratedSnakeCaseContract()
    {
        var setJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "set-document-property",
            ["session_id"] = _sessionId,
            ["property_name"] = "AutomationTag",
            ["value"] = "mcp-value",
            ["scope"] = "custom"
        });
        AssertSuccess(setJson, "workbook set-document-property");

        var builtInSetJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "set-document-property",
            ["session_id"] = _sessionId,
            ["property_name"] = "Title",
            ["value"] = "MCP workbook title",
            ["scope"] = "built-in"
        });
        AssertSuccess(builtInSetJson, "workbook set-document-property built-in");

        var getJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "get-document-property",
            ["session_id"] = _sessionId,
            ["property_name"] = "AutomationTag",
            ["scope"] = "custom"
        });

        AssertSuccess(getJson, "workbook get-document-property");
        using var document = JsonDocument.Parse(getJson);
        var property = document.RootElement.GetProperty("property");
        Assert.Equal("AutomationTag", property.GetProperty("name").GetString());
        Assert.Equal("mcp-value", property.GetProperty("value").GetString());
        Assert.Equal("custom", property.GetProperty("scope").GetString());

        var builtInGetJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "get-document-property",
            ["session_id"] = _sessionId,
            ["property_name"] = "Title",
            ["scope"] = "built-in"
        });
        AssertSuccess(builtInGetJson, "workbook get-document-property built-in");
        using var builtInDocument = JsonDocument.Parse(builtInGetJson);
        Assert.Equal(
            "MCP workbook title",
            builtInDocument.RootElement.GetProperty("property").GetProperty("value").GetString());

        var listJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "list-document-properties",
            ["session_id"] = _sessionId,
            ["include_built_in"] = true,
            ["include_custom"] = true
        });
        AssertSuccess(listJson, "workbook list-document-properties");
        using var listDocument = JsonDocument.Parse(listJson);
        var properties = listDocument.RootElement.GetProperty("properties").EnumerateArray().ToList();
        Assert.Contains(properties, item =>
            item.GetProperty("name").GetString() == "AutomationTag" &&
            item.GetProperty("scope").GetString() == "custom");
        Assert.Contains(properties, item =>
            item.GetProperty("name").GetString() == "Title" &&
            item.GetProperty("scope").GetString() == "built-in");

        var deleteJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "delete-document-property",
            ["session_id"] = _sessionId,
            ["property_name"] = "AutomationTag"
        });
        AssertSuccess(deleteJson, "workbook delete-document-property");

        var customListJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "list-document-properties",
            ["session_id"] = _sessionId,
            ["include_built_in"] = false,
            ["include_custom"] = true
        });
        AssertSuccess(customListJson, "workbook list-document-properties custom");
        using var customListDocument = JsonDocument.Parse(customListJson);
        Assert.DoesNotContain(
            customListDocument.RootElement.GetProperty("properties").EnumerateArray(),
            item => item.GetProperty("name").GetString() == "AutomationTag");
    }

    [Fact]
    public async Task SaveCopyAndExportFixedFormat_CreateExpectedFiles()
    {
        var copyPath = Path.Join(_tempDirectory, "WorkbookCopy.xlsx");
        var pdfPath = Path.Join(_tempDirectory, "WorkbookExport.pdf");
        var xpsPath = Path.Join(_tempDirectory, "WorkbookExport.xps");
        await SetCellValueAsync(_sessionId, 1);

        var copyJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "save-copy-as",
            ["session_id"] = _sessionId,
            ["target_path"] = copyPath
        });
        var exportJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "export-fixed-format",
            ["session_id"] = _sessionId,
            ["target_path"] = pdfPath,
            ["format_type"] = "pdf",
            ["quality"] = "standard",
            ["include_document_properties"] = true,
            ["ignore_print_areas"] = false,
            ["open_after_publish"] = false
        });
        var xpsJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "export-fixed-format",
            ["session_id"] = _sessionId,
            ["target_path"] = xpsPath,
            ["format_type"] = "xps",
            ["quality"] = "minimum",
            ["open_after_publish"] = false
        });

        AssertSuccess(copyJson, "workbook save-copy-as");
        AssertSuccess(exportJson, "workbook export-fixed-format");
        AssertSuccess(xpsJson, "workbook export-fixed-format xps");
        Assert.True(File.Exists(copyPath));
        Assert.True(File.Exists(pdfPath));
        Assert.True(File.Exists(xpsPath));
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsm")]
    [InlineData("xlsb")]
    [InlineData("xls")]
    public async Task SaveAs_UpdatesWorkbookAndSessionPaths(string format)
    {
        var outputPath = Path.Join(_tempDirectory, $"WorkbookSavedAs.{format}");

        var saveAsJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "save-as",
            ["session_id"] = _sessionId,
            ["target_path"] = outputPath,
            ["format"] = format
        });
        var infoJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "get-info",
            ["session_id"] = _sessionId
        });
        var sessionsJson = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "list"
        });

        AssertSuccess(saveAsJson, "workbook save-as");
        AssertSuccess(infoJson, "workbook get-info");
        using var infoDocument = JsonDocument.Parse(infoJson);
        Assert.Equal(Path.GetFullPath(outputPath), infoDocument.RootElement.GetProperty("fullName").GetString(), ignoreCase: true);
        using var sessionsDocument = JsonDocument.Parse(sessionsJson);
        var session = Assert.Single(sessionsDocument.RootElement.GetProperty("sessions").EnumerateArray());
        Assert.Equal(_sessionId, session.GetProperty("sessionId").GetString());
        Assert.Equal(Path.GetFullPath(outputPath), session.GetProperty("filePath").GetString(), ignoreCase: true);
    }

    [Fact]
    public async Task ExternalLinks_ListUpdateAndBreak_UsesGeneratedWorkbookRoutes()
    {
        var sourcePath = Path.Join(_tempDirectory, "LinkSource.xlsx");
        await CloseSessionAsync(_sessionId, save: true);

        var sourceSessionId = await CreateWorkbookSessionAsync(sourcePath);
        await SetCellValueAsync(sourceSessionId, 10);
        await CloseSessionAsync(sourceSessionId, save: true);

        _sessionId = await OpenWorkbookSessionAsync(_workbookPath);
        var sourceDirectory = Path.GetDirectoryName(sourcePath)!.Replace("'", "''", StringComparison.Ordinal);
        var formula = $"='{sourceDirectory}\\[{Path.GetFileName(sourcePath)}]Sheet1'!$A$1";
        await SetCellFormulaAsync(_sessionId, formula);
        await CloseSessionAsync(_sessionId, save: true);

        sourceSessionId = await OpenWorkbookSessionAsync(sourcePath);
        await SetCellValueAsync(sourceSessionId, 42);
        await CloseSessionAsync(sourceSessionId, save: true);

        _sessionId = await OpenWorkbookSessionAsync(_workbookPath);
        var listJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "list-external-links",
            ["session_id"] = _sessionId
        });
        AssertSuccess(listJson, "workbook list-external-links");
        using var listDocument = JsonDocument.Parse(listJson);
        var linkSource = Assert.Single(listDocument.RootElement.GetProperty("links").EnumerateArray())
            .GetProperty("source")
            .GetString();
        Assert.False(string.IsNullOrWhiteSpace(linkSource));

        var updateJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "update-external-link",
            ["session_id"] = _sessionId,
            ["link_source"] = linkSource
        });
        AssertSuccess(updateJson, "workbook update-external-link");

        var breakJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "break-external-link",
            ["session_id"] = _sessionId,
            ["link_source"] = linkSource
        });
        AssertSuccess(breakJson, "workbook break-external-link");

        var linksAfterBreakJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "list-external-links",
            ["session_id"] = _sessionId
        });
        AssertSuccess(linksAfterBreakJson, "workbook list-external-links after break");
        using var linksAfterBreakDocument = JsonDocument.Parse(linksAfterBreakJson);
        Assert.Empty(linksAfterBreakDocument.RootElement.GetProperty("links").EnumerateArray());

        var valuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1"
        });
        AssertSuccess(valuesJson, "range get-values after break");
        using var valuesDocument = JsonDocument.Parse(valuesJson);
        Assert.Equal(42d, valuesDocument.RootElement.GetProperty("values")[0][0].GetDouble());
    }

    private async Task<string> OpenWorkbookSessionAsync(string workbookPath)
    {
        var openJson = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["path"] = workbookPath
        });
        AssertSetupSuccess(openJson, $"file.open ({Path.GetFileName(workbookPath)})");

        using var openDocument = JsonDocument.Parse(openJson);
        var sessionId = openDocument.RootElement.GetProperty("session_id").GetString();
        TrackSession(sessionId);
        return Assert.IsType<string>(sessionId);
    }

    private async Task SetCellValueAsync(string sessionId, double value)
    {
        var resultJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["values"] = new List<List<object?>> { new() { value } }
        });
        AssertSetupSuccess(resultJson, "range.set-values");
    }

    private async Task SetCellFormulaAsync(string sessionId, string formula)
    {
        var resultJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-formulas",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["formulas"] = new List<List<string>> { new() { formula } }
        });
        AssertSetupSuccess(resultJson, "range.set-formulas");
    }
}
