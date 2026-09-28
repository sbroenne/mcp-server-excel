using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacPowerQueryRouteSelectorTests
{
    [Theory]
    [InlineData("list", null)]
    [InlineData("view", """{"queryName":"Sales"}""")]
    [InlineData("get-load-config", """{"queryName":"Sales"}""")]
    [InlineData("update", """{"queryName":"Sales","mCode":"let Source = 1 in Source","refresh":false}""")]
    public void SavedPackageActionsNeverRequireHelper(string action, string? args)
    {
        var route = MacPowerQueryRouteSelector.Select(action, Parse(args), Actions());

        Assert.Equal(MacPowerQueryRouteKind.SavedPackage, route.Kind);
        Assert.Null(route.HelperAction);
    }

    [Fact]
    public void UpdateWithDefaultRefreshSelectsHelperBeforeMutation()
    {
        var route = MacPowerQueryRouteSelector.Select(
            "update",
            Parse("""{"queryName":"Sales","mCode":"let Source = 1 in Source"}"""),
            Actions("powerquery.update"));

        Assert.Equal(MacPowerQueryRouteKind.Helper, route.Kind);
        Assert.Equal("powerquery.update", route.HelperAction);
        Assert.Equal("Sales", route.HelperArguments!["name"]!.GetValue<string>());
        Assert.Equal("let Source = 1 in Source", route.HelperArguments["formula"]!.GetValue<string>());
        Assert.True(route.HelperArguments["refresh"]!.GetValue<bool>());
    }

    [Fact]
    public void UpdateWithRefreshDoesNotFallBackWhenHelperMethodIsUnavailable()
    {
        var route = MacPowerQueryRouteSelector.Select(
            "update",
            Parse("""{"queryName":"Sales","mCode":"let Source = 1 in Source","refresh":true}"""),
            Actions());

        Assert.Equal(MacPowerQueryRouteKind.Unsupported, route.Kind);
        Assert.Contains("refresh", route.UnavailableReason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteRequestsExistingCleanupContractExplicitly()
    {
        var route = MacPowerQueryRouteSelector.Select(
            "delete",
            Parse("""{"queryName":"Sales"}"""),
            Actions("powerquery.delete"));

        Assert.Equal(MacPowerQueryRouteKind.Helper, route.Kind);
        Assert.Equal("Sales", route.HelperArguments!["name"]!.GetValue<string>());
        Assert.True(route.HelperArguments["deleteConnection"]!.GetValue<bool>());
    }

    [Fact]
    public void RenameUsesNormalizedNames()
    {
        var route = MacPowerQueryRouteSelector.Select(
            "rename",
            Parse("""{"oldName":" Sales ","newName":" Revenue "}"""),
            Actions("powerquery.rename"));

        Assert.Equal(MacPowerQueryRouteKind.Helper, route.Kind);
        Assert.Equal("Sales", route.HelperArguments!["name"]!.GetValue<string>());
        Assert.Equal("Revenue", route.HelperArguments["newName"]!.GetValue<string>());
    }

    [Fact]
    public void CreateUsesAtomicWorksheetDefaults()
    {
        var route = MacPowerQueryRouteSelector.Select(
            "create",
            Parse("""{"queryName":"Sales","mCode":"let Source = 1 in Source"}"""),
            Actions("powerquery.create"));

        Assert.Equal(MacPowerQueryRouteKind.Helper, route.Kind);
        Assert.Equal("load-to-table", route.HelperArguments!["destination"]!.GetValue<string>());
        Assert.Equal("Sales", route.HelperArguments["sheetName"]!.GetValue<string>());
        Assert.Equal("A1", route.HelperArguments["cellAddress"]!.GetValue<string>());
    }

    [Theory]
    [InlineData("load-to-data-model")]
    [InlineData("load-to-both")]
    public void CreateDataModelDestinationsRemainSeparatelyGated(string loadDestination)
    {
        var arguments = new JsonObject
        {
            ["queryName"] = "Sales",
            ["mCode"] = "let Source = 1 in Source"
        };
        if (loadDestination is not null)
        {
            arguments["loadDestination"] = loadDestination;
        }

        var route = MacPowerQueryRouteSelector.Select(
            "create",
            arguments,
            Actions("powerquery.create"));

        Assert.Equal(MacPowerQueryRouteKind.Unsupported, route.Kind);
        Assert.Contains("Data Model", route.UnavailableReason, StringComparison.Ordinal);
    }

    [Fact]
    public void CreateConnectionOnlyUsesHelperWithoutInventingARefresh()
    {
        var route = MacPowerQueryRouteSelector.Select(
            "create",
            Parse(
                """{"queryName":"Sales","mCode":"let Source = 1 in Source","loadDestination":"connection-only"}"""),
            Actions("powerquery.create"));

        Assert.Equal(MacPowerQueryRouteKind.Helper, route.Kind);
        Assert.Equal("Sales", route.HelperArguments!["name"]!.GetValue<string>());
        Assert.Equal("let Source = 1 in Source", route.HelperArguments["formula"]!.GetValue<string>());
        Assert.Equal("connection-only", route.HelperArguments["destination"]!.GetValue<string>());
        Assert.True(route.HelperArguments.ContainsKey("sheetName"));
        Assert.True(route.HelperArguments.ContainsKey("cellAddress"));
        Assert.Null(route.HelperArguments["sheetName"]);
        Assert.Null(route.HelperArguments["cellAddress"]);
    }

    [Fact]
    public void LoadToWorksheetMapsExactHelperArguments()
    {
        var route = MacPowerQueryRouteSelector.Select(
            "load-to",
            Parse(
                """{"queryName":"Sales","loadDestination":"load-to-table","targetSheet":"Report","targetCellAddress":"B5"}"""),
            Actions("powerquery.load-to"));

        Assert.Equal(MacPowerQueryRouteKind.Helper, route.Kind);
        Assert.Equal("Sales", route.HelperArguments!["name"]!.GetValue<string>());
        Assert.Equal("load-to-table", route.HelperArguments["destination"]!.GetValue<string>());
        Assert.Equal("Report", route.HelperArguments["sheetName"]!.GetValue<string>());
        Assert.Equal("B5", route.HelperArguments["cellAddress"]!.GetValue<string>());
    }

    [Fact]
    public void LoadToWorksheetUsesPublicDefaults()
    {
        var route = MacPowerQueryRouteSelector.Select(
            "load-to",
            Parse("""{"queryName":"Sales","loadDestination":"worksheet"}"""),
            Actions("powerquery.load-to"));

        Assert.Equal(MacPowerQueryRouteKind.Helper, route.Kind);
        Assert.Equal("load-to-table", route.HelperArguments!["destination"]!.GetValue<string>());
        Assert.Equal("Sales", route.HelperArguments["sheetName"]!.GetValue<string>());
        Assert.Equal("A1", route.HelperArguments["cellAddress"]!.GetValue<string>());
    }

    [Fact]
    public void ConnectionOnlyRejectsWorksheetCell()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            MacPowerQueryRouteSelector.Select(
                "load-to",
                Parse(
                    """{"queryName":"Sales","loadDestination":"connection-only","targetCellAddress":"B5"}"""),
                Actions("powerquery.load-to")));

        Assert.Contains("targetCellAddress", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LoadToConnectionOnlyUsesExplicitNullWorksheetArguments()
    {
        var route = MacPowerQueryRouteSelector.Select(
            "load-to",
            Parse("""{"queryName":"Sales","loadDestination":"connection-only"}"""),
            Actions("powerquery.load-to"));

        Assert.Equal(MacPowerQueryRouteKind.Helper, route.Kind);
        Assert.Equal("Sales", route.HelperArguments!["name"]!.GetValue<string>());
        Assert.Equal("connection-only", route.HelperArguments["destination"]!.GetValue<string>());
        Assert.True(route.HelperArguments.ContainsKey("sheetName"));
        Assert.True(route.HelperArguments.ContainsKey("cellAddress"));
        Assert.Null(route.HelperArguments["sheetName"]);
        Assert.Null(route.HelperArguments["cellAddress"]);
    }

    [Theory]
    [InlineData("load-to-data-model")]
    [InlineData("load-to-both")]
    public void DataModelLoadsRemainSeparatelyGated(string destination)
    {
        var route = MacPowerQueryRouteSelector.Select(
            "load-to",
            Parse($$"""{"queryName":"Sales","loadDestination":"{{destination}}"}"""),
            Actions("powerquery.load-to"));

        Assert.Equal(MacPowerQueryRouteKind.Unsupported, route.Kind);
        Assert.Contains("Data Model", route.UnavailableReason, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("refresh", """{"queryName":"Sales"}""", "powerquery.refresh")]
    [InlineData("refresh-all", "{}", "powerquery.refresh-all")]
    [InlineData("unload", """{"queryName":"Sales"}""", "powerquery.unload")]
    [InlineData("evaluate", """{"mCode":"let Source = 1 in Source"}""", "powerquery.evaluate")]
    public void EngineActionsUseOnlyAdvertisedHelperMethods(
        string action,
        string args,
        string helperAction)
    {
        var available = MacPowerQueryRouteSelector.Select(
            action,
            Parse(args),
            Actions(helperAction));
        var unavailable = MacPowerQueryRouteSelector.Select(
            action,
            Parse(args),
            Actions());

        Assert.Equal(MacPowerQueryRouteKind.Helper, available.Kind);
        Assert.Equal(helperAction, available.HelperAction);
        Assert.Equal(MacPowerQueryRouteKind.Unsupported, unavailable.Kind);
    }

    private static JsonObject Parse(string? json) =>
        json is null ? new JsonObject() : JsonNode.Parse(json)!.AsObject();

    private static HashSet<string> Actions(params string[] actions) =>
        new(actions, StringComparer.Ordinal);
}
