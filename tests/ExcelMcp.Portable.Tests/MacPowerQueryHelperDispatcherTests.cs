using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacPowerQueryHelperDispatcherTests
{
    [Fact]
    public async Task ListMapsCompletePublicMetadata()
    {
        var dispatcher = new MacPowerQueryHelperDispatcher((path, action, arguments, timeout) =>
            Task.FromResult(JsonSerializer.SerializeToElement(new
            {
                queries = new[]
                {
                    new
                    {
                        name = "Sales",
                        formulaPreview = "let Source = 1 in Source",
                        characterCount = 24,
                        loadMode = "load-to-table",
                        targetSheet = "Report",
                        isConnectionOnly = false,
                        isLoadedToDataModel = false
                    }
                }
            })));
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            "powerquery.list",
            new JsonObject());

        var result = await dispatcher.DispatchAsync(
            route,
            "/tmp/exact.xlsx",
            TimeSpan.FromSeconds(5),
            "list",
            new JsonObject());

        var query = Assert.Single(result.GetProperty("queries").EnumerateArray());
        Assert.True(result.GetProperty("success").GetBoolean());
        Assert.Equal("/tmp/exact.xlsx", result.GetProperty("filePath").GetString());
        Assert.Equal("Sales", query.GetProperty("name").GetString());
        Assert.Equal("load-to-table", query.GetProperty("loadMode").GetString());
        Assert.Equal("Report", query.GetProperty("targetSheet").GetString());
        Assert.False(query.TryGetProperty("formula", out _));
    }

    [Theory]
    [InlineData("view")]
    [InlineData("get-load-config")]
    public async Task ViewHelperMapsReadResultWithoutPackageInspection(string publicAction)
    {
        var dispatcher = new MacPowerQueryHelperDispatcher((path, action, arguments, timeout) =>
            Task.FromResult(JsonSerializer.SerializeToElement(new
            {
                name = "Sales",
                formula = "let Source = 1 in Source",
                characterCount = 24,
                loadMode = "connection-only",
                isConnectionOnly = true,
                isLoadedToDataModel = false,
                hasConnection = false
            })));
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            "powerquery.view",
            new JsonObject { ["name"] = "Sales" });

        var result = await dispatcher.DispatchAsync(
            route,
            "/tmp/exact.xlsx",
            TimeSpan.FromSeconds(5),
            publicAction,
            Parse("""{"queryName":"Sales"}"""));

        Assert.True(result.GetProperty("success").GetBoolean());
        Assert.Equal("Sales", result.GetProperty("queryName").GetString());
        Assert.Equal("connection-only", result.GetProperty("loadMode").GetString());
        Assert.False(result.TryGetProperty("targetSheet", out _));
        Assert.Equal(
            publicAction == "view",
            result.TryGetProperty("mCode", out _));
    }

    [Fact]
    public async Task ReadRejectsIncompleteHelperMetadata()
    {
        var dispatcher = new MacPowerQueryHelperDispatcher((path, action, arguments, timeout) =>
            Task.FromResult(JsonSerializer.SerializeToElement(new
            {
                name = "Sales",
                formula = "let Source = 1 in Source"
            })));
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            "powerquery.view",
            new JsonObject { ["name"] = "Sales" });

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            dispatcher.DispatchAsync(
                route,
                "/tmp/exact.xlsx",
                TimeSpan.FromSeconds(5),
                "view",
                Parse("""{"queryName":"Sales"}""")));
    }

    [Fact]
    public async Task MutationFailureIsNotRetriedThroughAnotherRoute()
    {
        var calls = 0;
        var dispatcher = new MacPowerQueryHelperDispatcher((path, action, arguments, timeout) =>
        {
            calls++;
            throw new TimeoutException("Mutation outcome is uncertain.");
        });
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            "powerquery.update",
            new JsonObject
            {
                ["name"] = "Sales",
                ["formula"] = "let Source = 1 in Source"
            });

        await Assert.ThrowsAsync<TimeoutException>(
            () => dispatcher.DispatchAsync(
                route,
                "/tmp/exact.xlsx",
                TimeSpan.FromSeconds(5),
                "update",
                Parse("""{"queryName":"Sales"}""")));

        Assert.Equal(1, calls);
    }

    [Fact]
    public async Task RenameMapsExistingPublicResultShape()
    {
        var dispatcher = new MacPowerQueryHelperDispatcher((path, action, arguments, timeout) =>
            Task.FromResult(JsonSerializer.SerializeToElement(new { })));
        var publicArguments = Parse("""{"oldName":" Sales ","newName":" Revenue "}""");
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            "powerquery.rename",
            new JsonObject
            {
                ["name"] = "Sales",
                ["newName"] = "Revenue"
            });

        var result = await dispatcher.DispatchAsync(
            route,
            "/tmp/exact.xlsx",
            TimeSpan.FromSeconds(5),
            "rename",
            publicArguments);

        Assert.True(result.GetProperty("success").GetBoolean());
        Assert.Equal("/tmp/exact.xlsx", result.GetProperty("filePath").GetString());
        Assert.Equal("power-query", result.GetProperty("objectType").GetString());
        Assert.Equal(" Sales ", result.GetProperty("oldName").GetString());
        Assert.Equal(" Revenue ", result.GetProperty("newName").GetString());
        Assert.Equal("Sales", result.GetProperty("normalizedOldName").GetString());
        Assert.Equal("Revenue", result.GetProperty("normalizedNewName").GetString());
    }

    [Theory]
    [InlineData("create")]
    [InlineData("update")]
    [InlineData("delete")]
    [InlineData("unload")]
    [InlineData("refresh-all")]
    public async Task OperationActionsReturnSuccessfulPublicResult(string action)
    {
        var dispatcher = new MacPowerQueryHelperDispatcher((path, helperAction, arguments, timeout) =>
            Task.FromResult(JsonSerializer.SerializeToElement(new { })));
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            $"powerquery.{action}",
            new JsonObject());

        var result = await dispatcher.DispatchAsync(
            route,
            "/tmp/exact.xlsx",
            TimeSpan.FromSeconds(5),
            action,
            new JsonObject());

        Assert.True(result.GetProperty("success").GetBoolean());
        Assert.Equal("/tmp/exact.xlsx", result.GetProperty("filePath").GetString());
    }

    [Fact]
    public async Task StructuredHelperResultRetainsActionFields()
    {
        var dispatcher = new MacPowerQueryHelperDispatcher((path, action, arguments, timeout) =>
            Task.FromResult(JsonSerializer.SerializeToElement(new
            {
                queryName = "Sales",
                hasErrors = false,
                errorMessages = Array.Empty<string>(),
                refreshTime = "2026-09-28T00:00:00Z",
                isConnectionOnly = false,
                loadedToSheet = "Report"
            })));
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            "powerquery.refresh",
            new JsonObject { ["name"] = "Sales" });

        var result = await dispatcher.DispatchAsync(
            route,
            "/tmp/exact.xlsx",
            TimeSpan.FromSeconds(5),
            "refresh",
            Parse("""{"queryName":"Sales"}"""));

        Assert.True(result.GetProperty("success").GetBoolean());
        Assert.Equal("Sales", result.GetProperty("queryName").GetString());
        Assert.Equal("Report", result.GetProperty("loadedToSheet").GetString());
    }

    [Fact]
    public async Task RefreshSuccessRejectsReportedQueryErrors()
    {
        var dispatcher = new MacPowerQueryHelperDispatcher((path, action, arguments, timeout) =>
            Task.FromResult(JsonSerializer.SerializeToElement(new
            {
                queryName = "Sales",
                hasErrors = true,
                errorMessages = new List<string> { "Formula.Firewall" },
                refreshTime = "2026-09-28T00:00:00Z",
                isConnectionOnly = false
            })));
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            "powerquery.refresh",
            new JsonObject { ["name"] = "Sales" });

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            dispatcher.DispatchAsync(
                route,
                "/tmp/exact.xlsx",
                TimeSpan.FromSeconds(5),
                "refresh",
                Parse("""{"queryName":"Sales"}""")));
    }

    [Fact]
    public async Task EvaluateMapsOnlyPublicTabularResult()
    {
        var dispatcher = new MacPowerQueryHelperDispatcher((path, action, arguments, timeout) =>
            Task.FromResult(JsonSerializer.SerializeToElement(new
            {
                columns = new List<string> { "Value" },
                rows = new List<List<object>> { new() { 42 } },
                rowCount = 1,
                columnCount = 1,
                temporaryQueryName = "__pq_eval_internal"
            })));
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            "powerquery.evaluate",
            new JsonObject { ["formula"] = "let Source = 42 in Source" });

        var result = await dispatcher.DispatchAsync(
            route,
            "/tmp/exact.xlsx",
            TimeSpan.FromSeconds(5),
            "evaluate",
            Parse("""{"mCode":"let Source = 42 in Source"}"""));

        Assert.True(result.GetProperty("success").GetBoolean());
        Assert.Equal("let Source = 42 in Source", result.GetProperty("mCode").GetString());
        Assert.Equal(1, result.GetProperty("rowCount").GetInt32());
        Assert.Equal(1, result.GetProperty("columnCount").GetInt32());
        Assert.False(result.TryGetProperty("temporaryQueryName", out _));
    }

    [Fact]
    public async Task EvaluateRejectsInconsistentResultDimensions()
    {
        var dispatcher = new MacPowerQueryHelperDispatcher((path, action, arguments, timeout) =>
            Task.FromResult(JsonSerializer.SerializeToElement(new
            {
                columns = new List<string> { "Value" },
                rows = new List<List<object>> { new() { 42 } },
                rowCount = 2,
                columnCount = 1
            })));
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            "powerquery.evaluate",
            new JsonObject { ["formula"] = "let Source = 42 in Source" });

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            dispatcher.DispatchAsync(
                route,
                "/tmp/exact.xlsx",
                TimeSpan.FromSeconds(5),
                "evaluate",
                Parse("""{"mCode":"let Source = 42 in Source"}""")));
    }

    [Fact]
    public async Task UnloadReturnsExistingPublicAction()
    {
        var dispatcher = new MacPowerQueryHelperDispatcher((path, action, arguments, timeout) =>
            Task.FromResult(JsonSerializer.SerializeToElement(new { removedTables = 1 })));
        var route = new MacPowerQueryRoute(
            MacPowerQueryRouteKind.Helper,
            "powerquery.unload",
            new JsonObject { ["name"] = "Sales" });

        var result = await dispatcher.DispatchAsync(
            route,
            "/tmp/exact.xlsx",
            TimeSpan.FromSeconds(5),
            "unload",
            Parse("""{"queryName":"Sales"}"""));

        Assert.Equal("unload", result.GetProperty("action").GetString());
        Assert.False(result.TryGetProperty("removedTables", out _));
    }

    private static JsonObject Parse(string json) => JsonNode.Parse(json)!.AsObject();
}
