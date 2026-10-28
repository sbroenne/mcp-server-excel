using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// CLI arguments and output for Power Query commands. Workbook behavior is
/// covered by the PersistentServicePowerQuery* Service tests.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PowerQueryContractCliTests
{
    [Theory]
    [InlineData("create")]
    [InlineData("update")]
    [InlineData("load-to")]
    [InlineData("refresh")]
    public async Task LoadMetrics_ArePreservedByCli(string action)
    {
        string[] args = action switch
        {
            "create" or "update" => ["--m-code", "#table({\"Value\"}, {{1}})"],
            "load-to" => ["--load-destination", "worksheet"],
            _ => []
        };
        var result = await InProcessCliHelper.RunAsync(
            ["powerquery", action, "--session", "session-1", "--query-name", "Query", .. args],
            _ => new ServiceResponse
            {
                Success = true,
                Result = """{"success":true,"queryName":"Query","durationMs":12.5,"rowsLoaded":1,"worksheetRowsLoaded":1}"""
            });
        Assert.Equal(0, result.ExitCode);
        using var document = JsonDocument.Parse(result.Stdout);
        Assert.Equal(12.5, document.RootElement.GetProperty("data").GetProperty("durationMs").GetDouble());
        Assert.Equal(1, document.RootElement.GetProperty("data").GetProperty("rowsLoaded").GetInt64());
        Assert.Equal(1, document.RootElement.GetProperty("data").GetProperty("worksheetRowsLoaded").GetInt64());
    }

    [Theory]
    [InlineData("create")]
    [InlineData("update")]
    public async Task Mutation_PrivacyFailure_ReportsCategoryAndPartialState(string action)
    {
        var partialState = JsonSerializer.SerializeToElement(new
        {
            queryName = "Query",
            queryExists = true,
            loadMode = "load-to-table",
            targetSheet = "Query",
            isLoadedToDataModel = false
        });
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["powerquery", action, "--session", "session-1", "--query-name", "Query", "--m-code", "error \"Formula.Firewall\""],
            request =>
            {
                captured = request;
                return new ServiceResponse
                {
                    Success = false,
                    Command = $"powerquery.{action}",
                    SessionId = "session-1",
                    ErrorMessage = "Formula.Firewall: query remains. To allow this combination, use powerquery set-privacy-settings --ignore-privacy-levels true.",
                    ErrorCategory = "PrivacyFirewall",
                    ExceptionType = "PowerQueryCommandException",
                    HResult = "0x800A03EC",
                    PartialState = partialState.GetRawText()
                };
            });

        Assert.Equal(1, result.ExitCode);
        Assert.NotNull(captured);
        Assert.Equal($"powerquery.{action}", captured.Command);
        using (var args = JsonDocument.Parse(captured.Args!))
        {
            Assert.Equal("Query", args.RootElement.GetProperty("queryName").GetString());
            Assert.Equal("error \"Formula.Firewall\"", args.RootElement.GetProperty("mCode").GetString());
        }
        using var document = JsonDocument.Parse(result.Stdout);
        var root = document.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("PrivacyFirewall", root.GetProperty("error").GetProperty("category").GetString());
        Assert.Equal("PowerQueryCommandException", root.GetProperty("error").GetProperty("exceptionType").GetString());
        Assert.Equal("0x800A03EC", root.GetProperty("error").GetProperty("hresult").GetString());
        Assert.Contains("powerquery set-privacy-settings", root.GetProperty("error").GetProperty("message").GetString(), StringComparison.Ordinal);
        Assert.Equal(partialState.GetRawText(), root.GetProperty("partialState").GetRawText());
    }

    [Fact]
    public async Task Refresh_PrivacyFailure_ReportsCategoryWithoutPartialState()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["powerquery", "refresh", "--session", "session-1", "--query-name", "Query", "--timeout-seconds", "60"],
            request =>
            {
                captured = request;
                return new ServiceResponse
                {
                    Success = false,
                    Command = "powerquery.refresh",
                    SessionId = "session-1",
                    ErrorMessage = "Formula.Firewall: Query 'ConfigData' references other queries.",
                    ErrorCategory = "PrivacyFirewall",
                    ExceptionType = "PowerQueryCommandException",
                    HResult = "0x800A03EC",
                    InnerError = "Formula.Firewall"
                };
            });

        Assert.Equal(1, result.ExitCode);
        Assert.NotNull(captured);
        using (var args = JsonDocument.Parse(captured.Args!))
        {
            Assert.Equal("Query", args.RootElement.GetProperty("queryName").GetString());
            Assert.Equal(60, args.RootElement.GetProperty("timeoutSeconds").GetInt32());
        }
        using var document = JsonDocument.Parse(result.Stdout);
        var root = document.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal("PrivacyFirewall", root.GetProperty("error").GetProperty("category").GetString());
        Assert.Equal("Formula.Firewall", root.GetProperty("error").GetProperty("innerError").GetString());
        Assert.False(root.TryGetProperty("partialState", out _));
    }

    [Fact]
    public async Task ListOmitsPreviewAndViewPreservesCompleteCode()
    {
        const string listJson = """{"success":true,"queries":[{"name":"Query","formulaPreview":"let","characterCount":5000,"loadMode":"connection-only"}]}""";
        const string viewJson = """{"success":true,"queryName":"Query","mCode":"let Source = 1 in Source","loadMode":"connection-only"}""";

        var list = await InProcessCliHelper.RunAsync(
            ["powerquery", "list", "--session", "session-1"],
            request =>
            {
                Assert.Equal("powerquery.list", request.Command);
                return new ServiceResponse { Success = true, Result = listJson };
            });
        var view = await InProcessCliHelper.RunAsync(
            ["powerquery", "view", "--session", "session-1", "--query-name", "Query"],
            request =>
            {
                Assert.Equal("powerquery.view", request.Command);
                using var args = JsonDocument.Parse(request.Args!);
                Assert.Equal("Query", args.RootElement.GetProperty("queryName").GetString());
                return new ServiceResponse { Success = true, Result = viewJson };
            });

        Assert.True(list.ExitCode == 0, list.Stdout + list.Stderr);
        Assert.True(view.ExitCode == 0, view.Stdout + view.Stderr);
        using var listDocument = JsonDocument.Parse(list.Stdout);
        var query = Assert.Single(listDocument.RootElement.GetProperty("data").GetProperty("queries").EnumerateArray());
        Assert.Equal("Query", query.GetProperty("name").GetString());
        Assert.Equal(5000, query.GetProperty("characterCount").GetInt32());
        Assert.Equal("connection-only", query.GetProperty("loadMode").GetString());
        Assert.False(query.TryGetProperty("formulaPreview", out _));
        using var viewDocument = JsonDocument.Parse(view.Stdout);
        using var expectedView = JsonDocument.Parse(viewJson);
        var viewData = viewDocument.RootElement.GetProperty("data");
        foreach (var property in expectedView.RootElement.EnumerateObject().Where(property => property.Name != "success"))
            Assert.True(JsonElement.DeepEquals(property.Value, viewData.GetProperty(property.Name)), property.Name);
        Assert.Equal(3, viewData.EnumerateObject().Count());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task PrivacySettings_MapSetAndGet(bool ignorePrivacyLevels)
    {
        var resultJson = $$"""{"success":true,"ignorePrivacyLevels":{{(ignorePrivacyLevels ? "true" : "false")}},"message":"global privacy option may override"}""";
        ServiceRequest? setRequest = null;
        var set = await InProcessCliHelper.RunAsync(
            ["powerquery", "set-privacy-settings", "--session", "session-1",
                "--ignore-privacy-levels", ignorePrivacyLevels ? "true" : "false"],
            request =>
            {
                setRequest = request;
                return new ServiceResponse { Success = true, Result = resultJson };
            });
        ServiceRequest? getRequest = null;
        var get = await InProcessCliHelper.RunAsync(
            ["powerquery", "get-privacy-settings", "--session", "session-1"],
            request =>
            {
                getRequest = request;
                return new ServiceResponse { Success = true, Result = resultJson };
            });

        Assert.True(set.ExitCode == 0, set.Stdout + set.Stderr);
        Assert.True(get.ExitCode == 0, get.Stdout + get.Stderr);
        Assert.NotNull(setRequest);
        Assert.Equal("powerquery.set-privacy-settings", setRequest.Command);
        using (var args = JsonDocument.Parse(setRequest.Args!))
            Assert.Equal(ignorePrivacyLevels, args.RootElement.GetProperty("ignorePrivacyLevels").GetBoolean());
        Assert.NotNull(getRequest);
        Assert.Equal("powerquery.get-privacy-settings", getRequest.Command);
        Assert.Equal("session-1", getRequest.SessionId);
        foreach (var output in new[] { set.Stdout, get.Stdout })
        {
            using var document = JsonDocument.Parse(output);
            Assert.Equal(ignorePrivacyLevels, document.RootElement.GetProperty("data").GetProperty("ignorePrivacyLevels").GetBoolean());
        }
    }
}
