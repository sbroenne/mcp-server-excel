using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeFindContractCliTests
{
    [Theory]
    [InlineData(null)]
    [InlineData(1)]
    [InlineData(25)]
    [InlineData(1001)]
    [InlineData(int.MaxValue)]
    public async Task Find_MapsLimitAndPreservesCoverage(int? maxMatches)
    {
        List<string> arguments =
        [
            "rangeedit", "find", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1:A26",
            "--search-value", "Apple", "--find-options", """{"matchEntireCell":true}"""
        ];
        if (maxMatches.HasValue)
        {
            arguments.Add("--max-matches");
            arguments.Add(maxMatches.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        var returnedCount = Math.Min(maxMatches ?? 10, 25);
        var response = new RangeFindResult
        {
            Success = true,
            TotalCount = 25,
            MatchingCells = Enumerable.Range(1, returnedCount)
                .Select(row => new RangeCell { Address = $"$A${row}", Row = row, Column = 1, Value = "Apple" })
                .ToList()
        };
        ServiceRequest? captured = null;

        var result = await InProcessCliHelper.RunAsync(arguments, request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(response, ServiceProtocol.JsonOptions)
            };
        });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeedit.find", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Sheet1", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("A1:A26", args.RootElement.GetProperty("rangeAddress").GetString());
        Assert.Equal("Apple", args.RootElement.GetProperty("searchValue").GetString());
        Assert.True(args.RootElement.GetProperty("findOptions").GetProperty("matchEntireCell").GetBoolean());
        if (maxMatches.HasValue)
        {
            Assert.Equal(maxMatches.Value, args.RootElement.GetProperty("maxMatches").GetInt32());
        }
        else
        {
            Assert.False(args.RootElement.TryGetProperty("maxMatches", out _));
        }
        using var document = JsonDocument.Parse(result.Stdout);
        var root = document.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean());
        Assert.Equal(25, root.GetProperty("totalCount").GetInt64());
        Assert.Equal(returnedCount, root.GetProperty("returnedCount").GetInt32());
        Assert.Equal(25 > returnedCount, root.GetProperty("truncated").GetBoolean());
        Assert.Equal(returnedCount, root.GetProperty("matchingCells").GetArrayLength());
    }

    [Theory]
    [InlineData("1.5")]
    [InlineData("2147483648")]
    public async Task Find_RejectsNonIntegerOrOverflowLimitBeforeDispatch(string limit)
    {
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "find", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1",
            "--search-value", "Apple", "--find-options", "{}",
            "--max-matches", limit
        ]);

        Assert.NotEqual(0, result.ExitCode);
        using var document = JsonDocument.Parse(result.Stdout);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains(limit, document.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("null")]
    [InlineData("[]")]
    [InlineData("{")]
    [InlineData("""{"matchEntireCell":"yes"}""")]
    public async Task Find_RejectsInvalidOptionsBeforeDispatch(string options)
    {
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "find", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1",
            "--search-value", "Apple", "--find-options", options
        ]);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("findOptions", result.Stdout + result.Stderr, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public async Task Find_PreservesNonpositiveLimitAndReportsServiceError(int maxMatches)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "find", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1",
            "--search-value", "Apple", "--find-options", "{}",
            "--max-matches", maxMatches.ToString(System.Globalization.CultureInfo.InvariantCulture)
        ], request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = "maxMatches must be positive.",
                ErrorCategory = "InvalidInput",
                ExceptionType = nameof(ArgumentOutOfRangeException)
            };
        });

        Assert.Equal(1, result.ExitCode);
        Assert.NotNull(captured);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(maxMatches, args.RootElement.GetProperty("maxMatches").GetInt32());
        using var document = JsonDocument.Parse(result.Stdout);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("maxMatches", document.RootElement.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Replace_MapsInheritedJsonOptions()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "replace", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1",
            "--find-value", "Apple", "--replace-value", "Pear",
            "--replace-options", """{"matchCase":true,"replaceAll":false}"""
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeedit.replace", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        var options = args.RootElement.GetProperty("replaceOptions");
        Assert.True(options.GetProperty("matchCase").GetBoolean());
        Assert.False(options.GetProperty("replaceAll").GetBoolean());
    }

    [Fact]
    public async Task Sort_RejectsFindOnlyLimitBeforeDispatch()
    {
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "sort", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1",
            "--sort-columns", """[{"columnIndex":1,"ascending":true}]""",
            "--max-matches", "1"
        ]);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains("maxMatches", result.Stdout + result.Stderr, StringComparison.Ordinal);
    }
}
