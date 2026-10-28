using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "Results")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ResultJsonTests
{
    [Fact]
    public void QueryList_NormalOutputOmitsPreviewAndKeepsCompleteMetadata()
    {
        var result = ResultJson.Read(
            """{"success":true,"queries":[{"name":"Source","formulaPreview":"let x = 1 in x","characterCount":14,"loadMode":"connection-only","isConnectionOnly":true}]}""",
            "powerquery.list");
        var query = Assert.Single(result.Data!["queries"]!.AsArray())!;
        Assert.Null(query["formulaPreview"]);
        Assert.Equal("Source", query["name"]!.GetValue<string>());
        Assert.Equal(14, query["characterCount"]!.GetValue<int>());
        Assert.Equal("connection-only", query["loadMode"]!.GetValue<string>());
        Assert.True(query["isConnectionOnly"]!.GetValue<bool>());
    }

    [Fact]
    public void QueryView_KeepsCompleteFormula()
    {
        var formula = "let source = \"" + new string('x', 500) + "\" in source";
        var result = ResultJson.Read(JsonSerializer.Serialize(new { success = true, mCode = formula }), "powerquery.view");
        Assert.Equal(formula, result.Data!["mCode"]!.GetValue<string>());
    }

    [Fact]
    public void ChartImageExport_KeepsDestinationInCliAndMcpResults()
    {
        const string command = "chart.export-image";
        const string destination = @"C:\exports\Revenue.png";
        var result = ResultJson.Read(
            """{"success":true,"filePath":"C:\\exports\\Revenue.png","action":"export-image"}""", command);
        using var cli = JsonDocument.Parse(result.ToCliJson(command, 1));
        using var mcp = JsonDocument.Parse(result.ToToolJson());
        Assert.Equal(destination, cli.RootElement.GetProperty("data").GetProperty("filePath").GetString());
        Assert.Equal(destination, mcp.RootElement.GetProperty("filePath").GetString());
    }

    [Fact]
    public void DictionaryResult_KeepsZeroAndNullEntries()
    {
        var result = ResultJson.Read(
            """{"success":true,"grandTotals":{"Revenue":0,"Returns":null}}""", "pivottablecalc.get-data");
        Assert.Equal("""{"Revenue":0,"Returns":null}""", result.Data!["grandTotals"]!.ToJsonString());
    }

    [Fact]
    public void TraceSummary_KeepsScopeUnreadableEntriesIssuesAndContinuation()
    {
        var result = ResultJson.Read(
            """{"success":true,"processId":42,"activityId":"activity-1","windowStartUtc":"2026-01-01T00:00:00Z","windowEndUtc":"2026-01-01T00:01:00Z","filesRead":2,"entriesMatched":10,"unreadableLines":3,"evaluations":4,"issues":[{"kind":"error","count":1}],"traceCursor":"2026-01-01T00:01:00Z"}""",
            "powerquerytrace.summarize");
        var data = result.Data!;
        Assert.Equal(42, data["processId"]!.GetValue<int>());
        Assert.Equal("activity-1", data["activityId"]!.GetValue<string>());
        Assert.Equal("2026-01-01T00:00:00Z", data["windowStartUtc"]!.GetValue<string>());
        Assert.Equal("2026-01-01T00:01:00Z", data["windowEndUtc"]!.GetValue<string>());
        Assert.Equal(2, data["filesRead"]!.GetValue<int>());
        Assert.Equal(10, data["entriesMatched"]!.GetValue<int>());
        Assert.Equal(3, data["unreadableLines"]!.GetValue<int>());
        Assert.Equal(4, data["evaluations"]!.GetValue<int>());
        Assert.Equal("error", data["issues"]![0]!["kind"]!.GetValue<string>());
        Assert.Equal(1, data["issues"]![0]!["count"]!.GetValue<int>());
        Assert.Equal("2026-01-01T00:01:00Z", data["traceCursor"]!.GetValue<string>());
    }

    [Fact]
    public void PageBreakResult_KeepsNativePrintScopeAndCoverage()
    {
        var result = ResultJson.Read(
            """{"success":true,"sheetName":"Data","printArea":"$A$1:$D$20","coverage":"Only breaks in the current print scope.","horizontal":[],"vertical":[]}""",
            "worksheetstyle.get-page-breaks");
        Assert.Equal("Data", result.Data!["sheetName"]!.GetValue<string>());
        Assert.Equal("$A$1:$D$20", result.Data["printArea"]!.GetValue<string>());
        Assert.Equal("Only breaks in the current print scope.", result.Data["coverage"]!.GetValue<string>());
        Assert.Empty(result.Data["horizontal"]!.AsArray());
        Assert.Empty(result.Data["vertical"]!.AsArray());
    }

    [Fact]
    public void ChartResult_KeepsUnavailableNativeGetterDetails()
    {
        var result = ResultJson.Read(
            """{"success":true,"fillTransparencyAvailable":false,"fillTransparencyReadError":"Marker transparency is not exposed.","markersSupported":true}""",
            "chartconfig.get-point-format");
        Assert.False(result.Data!["fillTransparencyAvailable"]!.GetValue<bool>());
        Assert.Equal("Marker transparency is not exposed.", result.Data["fillTransparencyReadError"]!.GetValue<string>());
        Assert.True(result.Data["markersSupported"]!.GetValue<bool>());
    }

    [Theory]
    [InlineData("unavailable", "Native window selection is unavailable: Excel rejected the read (HRESULT 0x800A03EC).")]
    [InlineData("unsupported", "The native selection does not expose a supported range/chart/shape interface.")]
    [InlineData("range", null)]
    [InlineData("chart", null)]
    [InlineData("shapes", null)]
    public void WindowResult_KeepsSelectionKindAndNativeDiagnostic(string selectionKind, string? selectionDiagnostic)
    {
        var address = selectionKind == "range" ? "$A$1" : null;
        var result = ResultJson.Read(JsonSerializer.Serialize(new
        {
            success = true,
            availability = "available",
            windows = new[] { new { selectionKind, selectionDiagnostic, selectedRangeAddress = address } }
        }), "window.get-context");
        Assert.True(result.Success);
        Assert.Equal("available", result.Data!["availability"]!.GetValue<string>());
        var window = Assert.Single(result.Data["windows"]!.AsArray())!;
        Assert.Equal(selectionKind, window["selectionKind"]!.GetValue<string>());
        Assert.Equal(selectionDiagnostic, window["selectionDiagnostic"]?.GetValue<string>());
        Assert.Equal(address, window["selectedRangeAddress"]?.GetValue<string>());
    }

    [Fact]
    public void WindowResult_KeepsNoWindowsAvailability()
    {
        var result = ResultJson.Read("""{"success":true,"availability":"no-windows","windows":[]}""", "window.get-context");
        Assert.True(result.Success);
        Assert.Equal("no-windows", result.Data!["availability"]!.GetValue<string>());
        Assert.Empty(result.Data["windows"]!.AsArray());
    }

    [Theory]
    [InlineData("rangeedit.get-filters")]
    [InlineData("tablecolumn.get-filters")]
    public void FilterResult_KeepsUnavailableCriterionAndNativeCause(string command)
    {
        var result = ResultJson.Read(
            """{"success":true,"columnFilters":[{"criteria1":{"available":false,"emptyVariant":false,"readError":"Native criterion is unavailable.","hResult":-2146827284}}]}""",
            command);
        var criterion = result.Data!["columnFilters"]![0]!["criteria1"]!;
        Assert.False(criterion["available"]!.GetValue<bool>());
        Assert.False(criterion["emptyVariant"]!.GetValue<bool>());
        Assert.Equal("Native criterion is unavailable.", criterion["readError"]!.GetValue<string>());
        Assert.Equal(-2146827284, criterion["hResult"]!.GetValue<int>());
    }

    [Fact]
    public void TableData_KeepsReturnedDimensionsAndValues()
    {
        var result = ResultJson.Read(
            """{"success":true,"data":[[false,0,null]],"rowCount":1,"columnCount":3}""", "table.get-data");
        Assert.Equal(1, result.Data!["rowCount"]!.GetValue<int>());
        Assert.Equal(3, result.Data["columnCount"]!.GetValue<int>());
        Assert.Equal("[[false,0,null]]", result.Data["data"]!.ToJsonString());
    }

    [Fact]
    public void FileTest_KeepsValidationAndAccessState()
    {
        var result = ResultJson.Read(
            """{"success":true,"canOpen":true,"isValid":true,"extension":".xlsx","message":"Read-only access","willOpenReadOnly":true}""",
            "session.test");
        Assert.True(result.Data!["canOpen"]!.GetValue<bool>());
        Assert.True(result.Data["isValid"]!.GetValue<bool>());
        Assert.True(result.Data["willOpenReadOnly"]!.GetValue<bool>());
        Assert.Equal("Read-only access", result.Data["message"]!.GetValue<string>());
    }

    [Fact]
    public void OptionalTraceFailure_DoesNotChangeCompletedRefresh()
    {
        var result = ResultJson.Read(
            """{"success":true,"queryName":"Report","traceSummary":{"success":false,"errorMessage":"Trace was unavailable"}}""",
            "powerquery.refresh", contract: ServiceRegistry.GetResultContract("powerquery.refresh"));
        Assert.True(result.Success);
        Assert.Null(result.Error);
        Assert.Equal("Trace was unavailable", result.Data!["traceSummary"]!["errorMessage"]!.GetValue<string>());
    }

    [Fact]
    public void NestedOperationFailure_PreservesCause()
    {
        var contract = ServiceRegistry.GetResultContract("powerquery.refresh-all");
        Assert.Contains("queryMetrics", contract!.OutcomePaths);
        var result = ResultJson.Read(
            """{"success":true,"queryMetrics":[{"success":false,"errorMessage":"Native refresh failed","errorCategory":"ComInterop","hresult":"0x800A03EC","innerError":"Provider unavailable","queryName":"Broken"}]}""",
            "powerquery.refresh-all", contract: contract);
        Assert.False(result.Success);
        Assert.Equal("Native refresh failed", result.Error!["message"]!.GetValue<string>());
        Assert.Equal("0x800A03EC", result.Error["hresult"]!.GetValue<string>());
        Assert.Equal("Provider unavailable", result.Error["innerError"]!.GetValue<string>());
        Assert.Equal("Broken", result.Data!["queryMetrics"]![0]!["queryName"]!.GetValue<string>());
    }

    [Theory]
    [InlineData("false")]
    [InlineData("0")]
    [InlineData("null")]
    [InlineData("\"0.00%\"")]
    [InlineData("[null,false,0]")]
    public void ScalarAndArrayResults_KeepTheirExactValues(string json)
    {
        using var output = JsonDocument.Parse(ResultJson.Read(json, "diag.read").ToCliJson("diag.read", 0));
        Assert.Equal(json, output.RootElement.GetProperty("data").GetRawText());
    }

    [Fact]
    public void VbaRead_KeepsPagingBoundaries()
    {
        var result = ResultJson.Read(
            """{"success":true,"code":"Sub Run()","startLine":1,"totalLineCount":12,"returnedLineCount":1,"hasMore":true,"nextStartLine":2,"moduleName":"Module1"}""",
            "vba.read");
        Assert.Equal("Sub Run()", result.Data!["code"]!.GetValue<string>());
        Assert.True(result.Data["hasMore"]!.GetValue<bool>());
        Assert.Equal(2, result.Data["nextStartLine"]!.GetValue<int>());
        Assert.Equal(12, result.Data["totalLineCount"]!.GetValue<int>());
        Assert.Equal(1, result.Data["returnedLineCount"]!.GetValue<int>());
    }

    [Fact]
    public void ExactRangeRead_PreservesGridAndDiagnostics()
    {
        var result = ResultJson.Read(
            """{"success":true,"filePath":"C:\\fixture.xlsx","sheetName":"Data","rangeAddress":"$A$1:$D$2","rowCount":2,"columnCount":4,"values":[[false,0,null,"#REF!"],[null,null,null,null]],"cellErrors":[{"cellAddress":"D1","errorName":"#REF!"}]}""",
            "range.get-values");
        using var output = JsonDocument.Parse(result.ToCliJson("range.get-values", 20));
        var data = output.RootElement.GetProperty("data");
        Assert.Equal("""[[false,0,null,"#REF!"],[null,null,null,null]]""", data.GetProperty("values").GetRawText());
        Assert.Equal(2, data.GetProperty("rowCount").GetInt32());
        Assert.Equal(4, data.GetProperty("columnCount").GetInt32());
        Assert.Equal("$A$1:$D$2", data.GetProperty("rangeAddress").GetString());
        Assert.Single(data.GetProperty("cellErrors").EnumerateArray());
    }

    [Fact]
    public void Failure_KeepsOriginalCauseWarningsAndPartialState()
    {
        var result = ResultJson.Read(
            """{"success":false,"errorMessage":"refresh failed","errorCategory":"Privacy","failedQueries":[{"queryName":"Report"}],"partialState":{"mCodeUpdated":true},"warnings":["stale data"]}""",
            "powerquery.refresh-all", contract: ServiceRegistry.GetResultContract("powerquery.refresh-all"));
        Assert.False(result.Success);
        Assert.Equal("Privacy", result.Error!["category"]!.GetValue<string>());
        Assert.True(result.PartialState!["mCodeUpdated"]!.GetValue<bool>());
        Assert.Equal("stale data", Assert.Single(result.Warnings!.AsArray())!.GetValue<string>());
        Assert.Equal("Report", Assert.Single(result.Data!["failedQueries"]!.AsArray())!["queryName"]!.GetValue<string>());
    }

    [Fact]
    public void BoundedInspection_KeepsDataScopeAndOmissions()
    {
        var result = ResultJson.Read(
            """{"success":true,"preview":{"sheetName":"Data","rangeAddress":"$A$1:$B$2","rowCount":2,"columnCount":2,"omittedRowCount":8,"omittedColumnCount":3,"truncatedTextCellCount":1,"values":[[1,2],[3,4]],"formulas":[["",""],["",""]]}}""",
            "workbook.inspect");
        Assert.Equal(8, result.Data!["preview"]!["omittedRowCount"]!.GetValue<int>());
        Assert.Equal(1, result.Data["preview"]!["truncatedTextCellCount"]!.GetValue<int>());
        Assert.Equal("[[1,2],[3,4]]", result.Data["preview"]!["values"]!.ToJsonString());
        Assert.Equal("""[["",""],["",""]]""", result.Data["preview"]!["formulas"]!.ToJsonString());
    }
}
