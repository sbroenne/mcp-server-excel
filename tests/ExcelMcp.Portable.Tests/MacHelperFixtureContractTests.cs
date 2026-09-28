using System.Text;
using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacHelperFixtureContractTests
{
    private const string RequestId = "0123456789abcdef0123456789abcdef";

    [Fact]
    public void RequestEnvelope_UsesExactVersionIdentityActionAndTypedArguments()
    {
        var workbookPath = Path.Combine(
            Path.GetTempPath(),
            "excelmcp-helper-fixture",
            "..",
            "fixture.xlsm");
        var request = MacHelperFixtureContract.CreateRequest(
            workbookPath,
            "helper.capabilities",
            new Dictionary<string, object?>(),
            RequestId);
        using var document = JsonDocument.Parse(request);
        var root = document.RootElement;

        Assert.Equal(1, root.GetProperty("version").GetInt32());
        Assert.Equal(RequestId, root.GetProperty("requestId").GetString());
        Assert.Equal(Path.GetFullPath(workbookPath), root.GetProperty("workbookPath").GetString());
        Assert.Equal("helper.capabilities", root.GetProperty("action").GetString());
        Assert.Empty(root.GetProperty("arguments").EnumerateObject());
        Assert.InRange(
            Encoding.UTF8.GetByteCount(request),
            1,
            MacHelperFixtureContract.MaximumUtf8Bytes);
    }

    [Theory]
    [InlineData("")]
    [InlineData("0123456789ABCDEF0123456789ABCDEF")]
    [InlineData("0123456789abcdef0123456789abcdeg")]
    public void RequestEnvelope_RejectsInvalidRequestIds(string requestId)
    {
        Assert.Throws<ArgumentException>(() => MacHelperFixtureContract.CreateRequest(
            Path.Combine(Path.GetTempPath(), "fixture.xlsm"),
            "helper.capabilities",
            new Dictionary<string, object?>(),
            requestId));
    }

    [Fact]
    public void PowerQueryLifecycle_UsesConfirmedConnectionOnlyContract()
    {
        Assert.Equal(
        [
            "powerquery.create",
            "powerquery.list",
            "powerquery.view",
            "powerquery.update",
            "powerquery.rename",
            "powerquery.view",
            "powerquery.delete",
            "powerquery.list"
        ],
            MacHelperFixtureContract.PowerQueryLifecycle.Select(step => step.Action));

        var create = MacHelperFixtureContract.PowerQueryLifecycle[0].Arguments;
        Assert.Equal(MacHelperFixtureContract.QueryName, create["name"]);
        Assert.Equal(MacHelperFixtureContract.QueryFormula, create["formula"]);
        var update = MacHelperFixtureContract.PowerQueryLifecycle[3].Arguments;
        Assert.Equal(MacHelperFixtureContract.UpdatedQueryFormula, update["formula"]);
        var rename = MacHelperFixtureContract.PowerQueryLifecycle[4].Arguments;
        Assert.Equal(MacHelperFixtureContract.RenamedQueryName, rename["newName"]);
        var delete = MacHelperFixtureContract.PowerQueryLifecycle[6].Arguments;
        Assert.Equal(true, delete["deleteConnection"]);
    }

    [Fact]
    public void VbaLifecycle_UsesExcelAuthoredWorkbookAndNeverRunsTheProcedure()
    {
        Assert.Equal(
        [
            "vba.import",
            "vba.list",
            "vba.view",
            "vba.update",
            "vba.view",
            "vba.delete",
            "vba.list"
        ],
            MacHelperFixtureContract.VbaLifecycle.Select(step => step.Action));
        Assert.DoesNotContain(
            MacHelperFixtureContract.VbaLifecycle,
            step => step.Action == "vba.run");

        var import = MacHelperFixtureContract.VbaLifecycle[0].Arguments;
        Assert.Equal(MacHelperFixtureContract.ModuleName, import["moduleName"]);
        Assert.Equal(MacHelperFixtureContract.ModuleSource, import["source"]);
        var update = MacHelperFixtureContract.VbaLifecycle[3].Arguments;
        Assert.Equal(MacHelperFixtureContract.UpdatedModuleSource, update["source"]);
    }

    [Fact]
    public void CapabilityResponse_KeepsPresenceSeparateFromProof()
    {
        var response = MacHelperFixtureContract.ParseResponse(
            $$"""
              {
                "version": 1,
                "requestId": "{{RequestId}}",
                "success": true,
                "result": {
                  "helperVersion": "1.0.1",
                  "protocolVersion": 1,
                  "staticAvailability": {
                    "queriesApi": true,
                    "queryTableApi": true,
                    "scenarioApi": true,
                    "vbProjectApi": true,
                    "codeModuleApi": true
                  },
                  "engineCapabilities": {
                    "xmlMapsApi": null,
                    "rangeXPathApi": null,
                    "workbookModelApi": null,
                    "dataModelConnectionApi": null
                  },
                  "supportedActions": [
                    "helper.capabilities",
                    "powerquery.list",
                    "powerquery.view",
                    "powerquery.create",
                    "powerquery.update",
                    "powerquery.rename",
                    "powerquery.delete",
                    "analysis.create-scenario",
                    "analysis.show-scenario",
                    "vba.list",
                    "vba.view",
                    "vba.import",
                    "vba.update",
                    "vba.delete"
                  ],
                  "trustReadiness": {
                    "powerQueryReadable": true,
                    "vbaProjectReadable": true
                  },
                  "provenMethods": {
                    "powerQueryList": false,
                    "powerQueryCreate": false,
                    "powerQueryUpdate": false,
                    "powerQueryRename": false,
                    "powerQueryDelete": false,
                    "powerQueryRefresh": false,
                    "powerQueryRefreshAll": false,
                    "powerQueryLoadTo": false,
                    "powerQueryUnload": false,
                    "powerQueryEvaluate": false,
                    "xmlXPathRead": false,
                    "dataModelRead": false,
                    "scenarioCreateShow": false,
                    "vbaListView": false,
                    "vbaMutation": false
                  }
                },
                "error": null
              }
              """,
            RequestId);
        var result = AssertSuccessEnvelope(response);

        Assert.Equal(MacHelperFixtureContract.HelperVersion, result.GetProperty("helperVersion").GetString());
        Assert.Equal(
            MacHelperFixtureContract.SupportedActions,
            result.GetProperty("supportedActions").EnumerateArray()
                .Select(action => action.GetString()!).ToArray());
        Assert.True(result.GetProperty("staticAvailability").GetProperty("queriesApi").GetBoolean());
        Assert.Equal(
            JsonValueKind.Null,
            result.GetProperty("engineCapabilities").GetProperty("workbookModelApi").ValueKind);
        Assert.True(result.GetProperty("trustReadiness").GetProperty("vbaProjectReadable").GetBoolean());
        Assert.All(
            result.GetProperty("provenMethods").EnumerateObject(),
            method => Assert.False(method.Value.GetBoolean()));
    }

    [Fact]
    public void FailureEnvelope_HasNoResultAndUsesStructuredSanitizedError()
    {
        var root = MacHelperFixtureContract.ParseResponse(
            $$"""
              {
                "version": 1,
                "requestId": "{{RequestId}}",
                "success": false,
                "result": null,
                "error": {
                  "category": "InvalidOperation",
                  "code": "helper_action_failed",
                  "message": "The requested helper action failed."
                }
              }
              """,
            RequestId);

        Assert.Equal(1, root.GetProperty("version").GetInt32());
        Assert.Equal(RequestId, root.GetProperty("requestId").GetString());
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal(JsonValueKind.Null, root.GetProperty("result").ValueKind);
        var error = root.GetProperty("error");
        Assert.Equal("InvalidOperation", error.GetProperty("category").GetString());
        Assert.Equal("helper_action_failed", error.GetProperty("code").GetString());
        Assert.Equal("The requested helper action failed.", error.GetProperty("message").GetString());
    }

    [Fact]
    public void LifecycleResultDtos_UseConfirmedNamesAndShapes()
    {
        var queryList = SuccessResult(
            """{"queries":[{"name":"ExcelMcpFixtureLiteral"}]}""");
        Assert.Equal(
            MacHelperFixtureContract.QueryName,
            Assert.Single(queryList.GetProperty("queries").EnumerateArray())
                .GetProperty("name").GetString());

        var queryView = SuccessResult(
            $$"""{"name":"{{MacHelperFixtureContract.QueryName}}","formula":{{JsonSerializer.Serialize(MacHelperFixtureContract.QueryFormula)}}}""");
        Assert.Equal(MacHelperFixtureContract.QueryFormula, queryView.GetProperty("formula").GetString());

        var moduleList = SuccessResult(
            """{"modules":[{"name":"ExcelMcpFixtureModule","type":1,"lineCount":5}]}""");
        var module = Assert.Single(moduleList.GetProperty("modules").EnumerateArray());
        Assert.Equal(MacHelperFixtureContract.ModuleName, module.GetProperty("name").GetString());
        Assert.Equal(1, module.GetProperty("type").GetInt32());
        Assert.Equal(5, module.GetProperty("lineCount").GetInt32());

        var moduleView = SuccessResult(
            $$"""{"moduleName":"{{MacHelperFixtureContract.ModuleName}}","moduleType":1,"lineCount":5,"source":{{JsonSerializer.Serialize(MacHelperFixtureContract.ModuleSource)}}}""");
        Assert.Equal(MacHelperFixtureContract.ModuleSource, moduleView.GetProperty("source").GetString());
    }

    [Fact]
    public void ResponseEnvelope_RejectsOversizedOrInconsistentResults()
    {
        var oversized = new string('x', MacHelperFixtureContract.MaximumUtf8Bytes + 1);
        Assert.Throws<InvalidDataException>(() =>
            MacHelperFixtureContract.ParseResponse(oversized, RequestId));
        Assert.Throws<InvalidDataException>(() =>
            MacHelperFixtureContract.ParseResponse(
                $$"""{"version":1,"requestId":"{{RequestId}}","success":true,"result":null,"error":null}""",
                RequestId));
    }

    private static JsonElement SuccessResult(string resultJson)
    {
        var response = MacHelperFixtureContract.ParseResponse(
            $$"""{"version":1,"requestId":"{{RequestId}}","success":true,"result":{{resultJson}},"error":null}""",
            RequestId);
        return AssertSuccessEnvelope(response);
    }

    private static JsonElement AssertSuccessEnvelope(JsonElement root)
    {
        Assert.Equal(1, root.GetProperty("version").GetInt32());
        Assert.Equal(RequestId, root.GetProperty("requestId").GetString());
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal(JsonValueKind.Null, root.GetProperty("error").ValueKind);
        return root.GetProperty("result");
    }
}
