// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class RangeMergedCellWriteRegressionTests : McpIntegrationTestBase
{
    private readonly string _testExcelFile;
    private string? _sessionId;

    public RangeMergedCellWriteRegressionTests(ITestOutputHelper output)
        : base(output, "RangeMergedCellWriteRegressionClient")
    {
        _testExcelFile = Path.Join(CreateTempDirectory("RangeMergedCellWrite"), "MergedCellWrite.xlsx");
    }

    protected override async Task InitializeTestAsync()
    {
        _sessionId = await CreateWorkbookSessionAsync(_testExcelFile);
    }

    [Fact]
    public async Task SetValues_MergedNonAnchorCell_ReturnsActionableFailureViaMcp()
    {
        AssertSuccess(
            await CallToolAsync("range", new Dictionary<string, object?>
            {
                ["action"] = "set-values",
                ["session_id"] = _sessionId,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1",
                ["values"] = new List<List<object?>> { new() { "Original" } }
            }),
            "range.set-values initial value");

        AssertSuccess(
            await CallToolAsync("range_format", new Dictionary<string, object?>
            {
                ["action"] = "merge-cells",
                ["session_id"] = _sessionId,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "A1:B1"
            }),
            "range_format.merge-cells");

        var failedWriteJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "B1",
            ["values"] = new List<List<object?>> { new() { "Updated" } }
        });

        using var failedWrite = ParseJsonResult(failedWriteJson, "range.set-values merged non-anchor");
        AssertFailureEnvelope(
            failedWrite.RootElement,
            "range.set-values merged non-anchor",
            nameof(InvalidOperationException));

        string? errorMessage = failedWrite.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("$A$1:$B$1", errorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("top-left", errorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("unmerge", errorMessage, StringComparison.OrdinalIgnoreCase);

        var readJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1"
        });

        using var read = JsonDocument.Parse(readJson);
        Assert.Equal("Original", read.RootElement.GetProperty("values")[0][0].GetString());
    }
}
