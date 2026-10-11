using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "VBA")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class VbaInspectionContractCliTests
{
    [Fact]
    public async Task Search_MapsOptionsAndPreservesLimitedMatchResults()
    {
        ServiceRequest? captured = null;
        var response = new VbaSearchResult
        {
            Success = true,
            HasMore = true,
            Matches = [new VbaSearchMatch { ModuleName = "Module1", Line = 4, Column = 12, Excerpt = "Debug.Print needle" }]
        };
        var result = await InProcessCliHelper.RunAsync(
            ["vba", "search", "--session", "session-1", "--search-text", "needle",
             "--module-name", "Module1", "--whole-word", "--match-case", "--max-matches", "1"],
            request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = JsonSerializer.Serialize(response, ServiceProtocol.JsonOptions) };
            });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("vba.search", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("needle", args.RootElement.GetProperty("searchText").GetString());
        Assert.Equal("Module1", args.RootElement.GetProperty("moduleName").GetString());
        Assert.True(args.RootElement.GetProperty("wholeWord").GetBoolean());
        Assert.True(args.RootElement.GetProperty("matchCase").GetBoolean());
        Assert.Equal(1, args.RootElement.GetProperty("maxMatches").GetInt32());
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.True(output.RootElement.GetProperty("hasMore").GetBoolean());
        var match = Assert.Single(output.RootElement.GetProperty("matches").EnumerateArray());
        Assert.Equal("Module1", match.GetProperty("moduleName").GetString());
        Assert.Equal(4, match.GetProperty("line").GetInt32());
        Assert.Equal(12, match.GetProperty("column").GetInt32());
    }

    [Theory]
    [InlineData("status")]
    [InlineData("references")]
    public async Task Inspection_MapsActionAndSession(string action)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["vba", action, "--session", "session-1"],
            request =>
            {
                captured = request;
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"vba.{action}", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
    }
}
