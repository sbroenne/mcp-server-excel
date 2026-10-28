using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PivotSharedCacheContractCliTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ShareExistingCache_MapsBothQualifiedReports(bool changed)
    {
        string response = $$"""{"success":true,"changed":{{(changed ? "true" : "false")}},"previousCacheIndex":{{(changed ? 3 : 2)}},"cacheIndex":2,"targetCacheIndex":2,"sharedPivotTables":[{"sheetName":"Overview","pivotTableName":"First"},{"sheetName":"Summary","pivotTableName":"Second"}],"previousCacheUsers":[]}""";
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["pivottable", "share-cache", "--session", "session-1", "--sheet-name", "Summary",
                "--pivot-table-name", "Second", "--share-cache-with", "First", "--share-cache-with-sheet", "Overview"],
            request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = response };
            });
        Assert.Equal(0, result.ExitCode);
        Assert.NotNull(captured);
        Assert.Equal("pivottable.share-cache", captured.Command);
        using var actual = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Summary", actual.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("Second", actual.RootElement.GetProperty("pivotTableName").GetString());
        Assert.Equal("First", actual.RootElement.GetProperty("shareCacheWith").GetString());
        Assert.Equal("Overview", actual.RootElement.GetProperty("shareCacheWithSheet").GetString());
        using var expected = JsonDocument.Parse(response);
        using var returned = JsonDocument.Parse(result.Stdout);
        var data = returned.RootElement.GetProperty("data");
        Assert.Equal(expected.RootElement.EnumerateObject().Count() - 1, data.EnumerateObject().Count());
        foreach (var property in expected.RootElement.EnumerateObject().Where(property => property.Name != "success"))
            Assert.True(JsonElement.DeepEquals(property.Value, data.GetProperty(property.Name)), property.Name);
    }

    [Theory]
    [InlineData("--sheet-name")]
    [InlineData("--pivot-table-name")]
    [InlineData("--share-cache-with")]
    [InlineData("--share-cache-with-sheet")]
    public async Task ShareExistingCache_MissingRequiredInputDoesNotDispatch(string missing)
    {
        string[] options = ["--sheet-name", "Summary", "--pivot-table-name", "Second",
            "--share-cache-with", "First", "--share-cache-with-sheet", "Overview"];
        var args = new List<string> { "pivottable", "share-cache", "--session", "session-1" };
        for (int i = 0; i < options.Length; i += 2)
            if (options[i] != missing)
                args.AddRange([options[i], options[i + 1]]);
        bool dispatched = false;
        var result = await InProcessCliHelper.RunAsync(args, _ =>
        {
            dispatched = true;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.NotEqual(0, result.ExitCode);
        Assert.False(dispatched);
    }

    [Fact]
    public async Task ShareExistingCache_BatchJsonPreservesQualifiedReports()
    {
        const string args = """{"sheetName":"Summary","pivotTableName":"Second","shareCacheWith":"First","shareCacheWithSheet":"Overview"}""";
        string input = $$"""{"command":"pivottable.share-cache","sessionId":"session-1","args":{{args}}}""";
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(["batch", "--input", "-"], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true,"changed":true}""" };
        }, input + Environment.NewLine);
        Assert.Equal(0, result.ExitCode);
        Assert.NotNull(captured);
        Assert.Equal("pivottable.share-cache", captured.Command);
        using var actual = JsonDocument.Parse(captured.Args!);
        using var expected = JsonDocument.Parse(args);
        Assert.True(JsonElement.DeepEquals(expected.RootElement, actual.RootElement));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task CreationFlags_MapSharingPairAndOmittedDefaults(bool table, bool share)
    {
        var action = table ? "create-from-table" : "create-from-range";
        var arguments = new List<string>
        {
            "pivottable", action, "--session", "session-1", "--pivot-table-name", "Second",
            "--destination-sheet", "Summary", "--destination-cell", "H1"
        };
        arguments.AddRange(table ? ["--table-name", "Sales"] : ["--source-sheet", "Data", "--source-range", "A1:C20"]);
        if (share)
            arguments.AddRange(["--share-cache-with", "First", "--share-cache-with-sheet", "Other Summary"]);
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(arguments, request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"pivottable.{action}", captured.Command);
        using var json = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Second", json.RootElement.GetProperty("pivotTableName").GetString());
        Assert.Equal("Summary", json.RootElement.GetProperty("destinationSheet").GetString());
        if (share)
        {
            Assert.Equal("First", json.RootElement.GetProperty("shareCacheWith").GetString());
            Assert.Equal("Other Summary", json.RootElement.GetProperty("shareCacheWithSheet").GetString());
        }
        else
        {
            Assert.False(json.RootElement.TryGetProperty("shareCacheWith", out _));
            Assert.False(json.RootElement.TryGetProperty("shareCacheWithSheet", out _));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BatchJson_PreservesCamelCaseSharingArguments(bool table)
    {
        var action = table ? "create-from-table" : "create-from-range";
        var args = new Dictionary<string, object?>
        {
            ["pivotTableName"] = "Second",
            ["destinationSheet"] = "Summary",
            ["destinationCell"] = "H1",
            ["shareCacheWith"] = "First",
            ["shareCacheWithSheet"] = "Other Summary"
        };
        if (table)
            args["tableName"] = "Sales";
        else
        {
            args["sourceSheet"] = "Data";
            args["sourceRange"] = "A1:C20";
        }
        var input = JsonSerializer.Serialize(new { command = $"pivottable.{action}", sessionId = "session-1", args });
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(["batch", "--input", "-"], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        }, input + Environment.NewLine);
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"pivottable.{action}", captured.Command);
        using var actual = JsonDocument.Parse(captured.Args!);
        Assert.True(JsonElement.DeepEquals(JsonSerializer.SerializeToElement(args), actual.RootElement));
    }

    [Fact]
    public async Task SharingFlags_AreNotApplicableToDataModelCreation()
    {
        var result = await InProcessCliHelper.RunWithServiceAsync(
            ["pivottable", "create-from-datamodel", "--session", "missing-session", "--pivot-table-name", "Second",
                "--table-name", "Sales", "--destination-sheet", "Summary", "--destination-cell", "H1",
                "--share-cache-with", "First", "--share-cache-with-sheet", "Summary"]);
        Assert.Equal(1, result.ExitCode);
        Assert.Contains("shareCacheWith", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("session", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
    }
}
