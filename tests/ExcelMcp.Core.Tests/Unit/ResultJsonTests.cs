using System.Text;
using System.Text.Json;
using System.Reflection;
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
    public void TraceSummarySelection_KeepsScopeUnreadableEntriesIssuesAndContinuation()
    {
        const string command = "powerquerytrace.summarize";
        var result = ResultJson.Read(
            """{"success":true,"processId":42,"activityId":"activity-1","windowStartUtc":"2026-01-01T00:00:00Z","windowEndUtc":"2026-01-01T00:01:00Z","filesRead":2,"entriesMatched":10,"unreadableLines":3,"evaluations":4,"issues":[{"kind":"error","count":1}],"traceCursor":"2026-01-01T00:01:00Z"}""",
            command);
        result.Select(command, ["evaluations"], ServiceRegistry.GetResultContract(command));
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
    public void PageBreakSelection_KeepsNativePrintScopeAndCoverage()
    {
        const string command = "worksheetstyle.get-page-breaks";
        var result = ResultJson.Read(
            """{"success":true,"sheetName":"Data","printArea":"$A$1:$D$20","coverage":"Only breaks in the current print scope.","horizontal":[],"vertical":[]}""",
            command);
        result.Select(command, ["horizontal"], ServiceRegistry.GetResultContract(command));
        Assert.Equal("Data", result.Data!["sheetName"]!.GetValue<string>());
        Assert.Equal("$A$1:$D$20", result.Data["printArea"]!.GetValue<string>());
        Assert.Equal("Only breaks in the current print scope.", result.Data["coverage"]!.GetValue<string>());
        Assert.Empty(result.Data["horizontal"]!.AsArray());
        Assert.Null(result.Data["vertical"]);
    }

    [Fact]
    public void ChartSelection_KeepsUnavailableNativeGetterDetails()
    {
        const string command = "chartconfig.get-point-format";
        var result = ResultJson.Read(
            """{"success":true,"fillTransparencyAvailable":false,"fillTransparencyReadError":"Marker transparency is not exposed.","markersSupported":true}""",
            command);
        result.Select(command, ["fillTransparency"], ServiceRegistry.GetResultContract(command));
        Assert.False(result.Data!["fillTransparencyAvailable"]!.GetValue<bool>());
        Assert.Equal("Marker transparency is not exposed.", result.Data["fillTransparencyReadError"]!.GetValue<string>());
        Assert.True(result.Data["markersSupported"]!.GetValue<bool>());
    }

    [Theory]
    [InlineData("rangeedit.get-filters")]
    [InlineData("tablecolumn.get-filters")]
    public void FilterSelection_KeepsUnavailableCriterionAndNativeCause(string command)
    {
        var result = ResultJson.Read(
            """{"success":true,"columnFilters":[{"criteria1":{"available":false,"emptyVariant":false,"readError":"Native criterion is unavailable.","hResult":-2146827284}}]}""",
            command);
        result.Select(command, ["columnFilters.criteria1.value"], ServiceRegistry.GetResultContract(command));
        var criterion = result.Data!["columnFilters"]![0]!["criteria1"]!;
        Assert.False(criterion["available"]!.GetValue<bool>());
        Assert.False(criterion["emptyVariant"]!.GetValue<bool>());
        Assert.Equal("Native criterion is unavailable.", criterion["readError"]!.GetValue<string>());
        Assert.Equal(-2146827284, criterion["hResult"]!.GetValue<int>());
    }

    [Fact]
    public void TableDataSelection_KeepsReturnedDimensions()
    {
        var result = ResultJson.Read(
            """{"success":true,"data":[[false,0,null]],"rowCount":1,"columnCount":3}""",
            "table.get-data");
        result.Select("table.get-data", ["data"], ServiceRegistry.GetResultContract("table.get-data"));
        Assert.Equal(1, result.Data!["rowCount"]!.GetValue<int>());
        Assert.Equal(3, result.Data["columnCount"]!.GetValue<int>());
    }

    [Fact]
    public void FileTestSelection_KeepsValidationAndAccessState()
    {
        var result = ResultJson.Read(
            """{"success":true,"canOpen":true,"isValid":true,"extension":".xlsx","message":"Read-only access","willOpenReadOnly":true}""",
            "session.test");
        result.Select("session.test", ["extension"], ServiceRegistry.GetResultContract("session.test"));
        Assert.True(result.Data!["canOpen"]!.GetValue<bool>());
        Assert.True(result.Data["isValid"]!.GetValue<bool>());
        Assert.True(result.Data["willOpenReadOnly"]!.GetValue<bool>());
        Assert.Equal("Read-only access", result.Data["message"]!.GetValue<string>());
    }

    [Fact]
    public void ScalarReadSelection_KeepsCompleteValue()
    {
        var contract = ServiceRegistry.GetResultContract("chartconfig.get-axis-number-format");
        Assert.NotNull(contract);
        ResultJson.ValidateFields("chartconfig.get-axis-number-format", ["result"], contract);
        var result = ResultJson.Read("\"0.00%\"", "chartconfig.get-axis-number-format");
        result.Select("chartconfig.get-axis-number-format", ["result"], contract);
        Assert.Equal("0.00%", result.Data!.GetValue<string>());
    }

    [Fact]
    public void EveryGeneratedRead_HasUsableSourceFieldsAndProtectedPaths()
    {
        var readCount = 0;
        foreach (var category in typeof(ServiceRegistry).GetNestedTypes(BindingFlags.Public))
        {
            if (category.GetField("Category")?.GetRawConstantValue() is not string categoryName ||
                category.GetField("ValidActions")?.GetValue(null) is not string[] actions)
                continue;
            foreach (var action in actions)
            {
                var command = categoryName + "." + action;
                var contract = ServiceRegistry.GetResultContract(command);
                Assert.NotNull(contract);
                if (!contract.ReadOnly)
                    continue;
                readCount++;
                Assert.True(contract.Fields.Length > 0, $"No selectable result fields for {command}.");
                Assert.All(contract.ContextFields, field => Assert.Contains(field, contract.Fields));
                Assert.All(contract.OutcomePaths, field => Assert.Contains(field, contract.Fields));
                foreach (var field in contract.Fields)
                    ResultJson.ValidateFields(command, [field], contract);
            }
        }
        Assert.True(readCount > 50);
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
    public void NestedOperationFailure_PreservesCauseAndCannotBeFiltered()
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
    [InlineData("[null,false,0]")]
    public void ScalarAndArrayResults_KeepTheirExactValues(string json)
    {
        using var output = JsonDocument.Parse(ResultJson.Read(json, "diag.read").ToCliJson("diag.read", 0));
        Assert.Equal(json, output.RootElement.GetProperty("data").GetRawText());
    }

    [Theory]
    [InlineData("""{"fields":null}""")]
    [InlineData("""{"fields":"queries.name"}""")]
    [InlineData("""{"fields":[1]}""")]
    [InlineData("[]")]
    public void InvalidFieldInputs_ReportExpectedShape(string input)
    {
        var exception = Assert.Throws<ArgumentException>(() => ResultJson.ReadFields(input));
        Assert.DoesNotContain("System.", exception.Message);
        Assert.Contains(input == "[]" ? "JSON object" : "array", exception.Message);
    }

    [Fact]
    public void VbaSelection_KeepsPagingBoundaries()
    {
        var result = ResultJson.Read(
            """{"success":true,"code":"Sub Run()","startLine":1,"totalLineCount":12,"returnedLineCount":1,"hasMore":true,"nextStartLine":2,"moduleName":"Module1"}""",
            "vba.read");
        result.Select("vba.read", ["code"], ServiceRegistry.GetResultContract("vba.read"));
        Assert.True(result.Data!["hasMore"]!.GetValue<bool>());
        Assert.Equal(2, result.Data["nextStartLine"]!.GetValue<int>());
        Assert.Equal(12, result.Data["totalLineCount"]!.GetValue<int>());
        Assert.Equal(1, result.Data["returnedLineCount"]!.GetValue<int>());
    }

    [Fact]
    public void QuerySelection_UsesSourceFieldsEvenForEmptyCollections()
    {
        var contract = ServiceRegistry.GetResultContract("powerquery.list");
        Assert.NotNull(contract);
        Assert.Contains("queries.name", contract.Fields);
        Assert.Contains("queries.formulaPreview", contract.Fields);
        var result = ResultJson.Read("""{"success":true,"queries":[]}""", "powerquery.list");
        result.Select("powerquery.list", ["queries.name"], contract);
        Assert.Equal("""{"queries":[]}""", result.Data!.ToJsonString());
        Assert.Throws<ArgumentException>(() =>
            result.Select("powerquery.list", ["queries.unknown"], contract));
    }

    [Fact]
    public void QueryList_DefaultOmitsPreview_ExplicitSelectionRetainsIt()
    {
        const string input = """{"success":true,"queries":[{"name":"Source","formulaPreview":"let x = 1 in x","loadMode":"connection-only","isConnectionOnly":true}]}""";
        var contract = ServiceRegistry.GetResultContract("powerquery.list");
        var compact = ResultJson.Read(input, "powerquery.list");
        compact.Select("powerquery.list", null, contract);
        Assert.Null(compact.Data!["queries"]![0]!["formulaPreview"]);
        var selected = ResultJson.Read(input, "powerquery.list");
        selected.Select("powerquery.list", ["queries.name", "queries.formulaPreview"], contract);
        Assert.Equal("let x = 1 in x", selected.Data!["queries"]![0]!["formulaPreview"]!.GetValue<string>());
        Assert.Null(selected.Data!["queries"]![0]!["loadMode"]);
        Assert.True(Encoding.UTF8.GetByteCount(compact.ToCliJson("powerquery.list", 1)) <
            Encoding.UTF8.GetByteCount(ResultJson.Read(input, "powerquery.list").ToCliJson("powerquery.list", 1)));
    }

    [Fact]
    public void ExactRangeSelection_PreservesGridAndDiagnostics()
    {
        const string input = """{"success":true,"filePath":"C:\\fixture.xlsx","sheetName":"Data","rangeAddress":"$A$1:$D$2","rowCount":2,"columnCount":4,"values":[[false,0,null,"#REF!"],[null,null,null,null]],"cellErrors":[{"cellAddress":"D1","errorName":"#REF!"}]}""";
        var result = ResultJson.Read(input, "range.get-values");
        result.Select("range.get-values", ["values"], ServiceRegistry.GetResultContract("range.get-values"));
        using var output = JsonDocument.Parse(result.ToCliJson("range.get-values", 20));
        var data = output.RootElement.GetProperty("data");
        Assert.Equal("""[[false,0,null,"#REF!"],[null,null,null,null]]""", data.GetProperty("values").GetRawText());
        Assert.Equal(2, data.GetProperty("rowCount").GetInt32());
        Assert.Equal(4, data.GetProperty("columnCount").GetInt32());
        Assert.Equal("$A$1:$D$2", data.GetProperty("rangeAddress").GetString());
        Assert.Single(data.GetProperty("cellErrors").EnumerateArray());
    }

    [Fact]
    public void FailureSelection_CannotHideOriginalCauseOrPartialState()
    {
        var result = ResultJson.Read(
            """{"success":false,"errorMessage":"refresh failed","errorCategory":"Privacy","failedQueries":[{"queryName":"Report"}],"partialState":{"mCodeUpdated":true},"warnings":["stale data"]}""",
            "powerquery.refresh-all", contract: ServiceRegistry.GetResultContract("powerquery.refresh-all"));
        result.Select("powerquery.refresh-all", ["refreshedQueries"], ServiceRegistry.GetResultContract("powerquery.refresh-all"));
        Assert.False(result.Success);
        Assert.Equal("Privacy", result.Error!["category"]!.GetValue<string>());
        Assert.True(result.PartialState!["mCodeUpdated"]!.GetValue<bool>());
        Assert.Single(result.Warnings!.AsArray());
        Assert.Single(result.Data!["failedQueries"]!.AsArray());
    }

    [Fact]
    public void BoundedInspectionSelection_RetainsScopeAndOmissions()
    {
        var result = ResultJson.Read(
            """{"success":true,"preview":{"sheetName":"Data","rangeAddress":"$A$1:$B$2","rowCount":2,"columnCount":2,"omittedRowCount":8,"omittedColumnCount":3,"truncatedTextCellCount":1,"values":[[1,2],[3,4]],"formulas":[["",""],["",""]]}}""",
            "workbook.inspect");
        result.Select("workbook.inspect", ["preview.values"], ServiceRegistry.GetResultContract("workbook.inspect"));
        Assert.Equal(8, result.Data!["preview"]!["omittedRowCount"]!.GetValue<int>());
        Assert.Equal(1, result.Data!["preview"]!["truncatedTextCellCount"]!.GetValue<int>());
        Assert.Null(result.Data!["preview"]!["formulas"]);
    }
}
