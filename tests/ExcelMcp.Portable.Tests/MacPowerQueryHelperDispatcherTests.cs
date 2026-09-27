using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacPowerQueryHelperDispatcherTests
{
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

    private static JsonObject Parse(string json) => JsonNode.Parse(json)!.AsObject();
}
