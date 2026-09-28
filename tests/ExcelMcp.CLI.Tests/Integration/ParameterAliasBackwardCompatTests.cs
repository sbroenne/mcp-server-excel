using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Real parser coverage for backward-compatible range parameter aliases.
/// Workbook behavior is owned by the corresponding Service range tests.
/// </summary>
[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ParameterAliasBackwardCompatTests
{
    [Fact]
    public async Task RangeSetValues_ShortAliases_MapToCanonicalRequest()
    {
        var (result, requests) = await InProcessCliHelper.RunRecordingAsync(
            "range set-values --session session-1 --sheet AliasSheet --range C7 --values \"[[\\\"BackwardCompat\\\"]]\"",
            new ServiceResponse
            {
                Success = true,
                Result = """{"success":true}"""
            });

        Assert.Equal(0, result.ExitCode);
        var request = Assert.Single(requests);
        Assert.Equal("range.set-values", request.Command);
        Assert.Equal("session-1", request.SessionId);
        using var args = JsonDocument.Parse(request.Args!);
        Assert.Equal("AliasSheet", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("C7", args.RootElement.GetProperty("rangeAddress").GetString());
        Assert.Equal(
            "BackwardCompat",
            args.RootElement.GetProperty("values")[0][0].GetString());

        using var output = JsonDocument.Parse(result.Stdout);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task RangeGetValues_ShortAliases_MapToCanonicalRequestAndResult()
    {
        var (result, requests) = await InProcessCliHelper.RunRecordingAsync(
            "range get-values --session session-1 --sheet AliasSheet --range E9",
            new ServiceResponse
            {
                Success = true,
                Result = """{"success":true,"values":[["ReadTest"]]}"""
            });

        Assert.Equal(0, result.ExitCode);
        var request = Assert.Single(requests);
        Assert.Equal("range.get-values", request.Command);
        Assert.Equal("session-1", request.SessionId);
        using var args = JsonDocument.Parse(request.Args!);
        Assert.Equal("AliasSheet", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("E9", args.RootElement.GetProperty("rangeAddress").GetString());

        using var output = JsonDocument.Parse(result.Stdout);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(
            "ReadTest",
            output.RootElement.GetProperty("values")[0][0].GetString());
    }
}
