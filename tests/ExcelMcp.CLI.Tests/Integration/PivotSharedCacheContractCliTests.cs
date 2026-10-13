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
