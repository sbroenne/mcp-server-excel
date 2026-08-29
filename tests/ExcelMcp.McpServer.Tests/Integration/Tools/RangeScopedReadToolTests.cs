using System.Diagnostics;
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
public sealed class RangeScopedReadToolTests : McpIntegrationTestBase
{
    private readonly string _tempDirectory;

    public RangeScopedReadToolTests(ITestOutputHelper output)
        : base(output, "RangeScopedReadToolClient")
    {
        _tempDirectory = CreateTempDirectory(nameof(RangeScopedReadToolTests));
    }

    [Fact]
    public async Task GetValues_WithScope_ReturnsPageAndMetadataThroughMcp()
    {
        var baselineExcelProcessIds = Process.GetProcessesByName("EXCEL")
            .Select(process => process.Id)
            .ToHashSet();
        var sessionId = await CreateWorkbookSessionAsync(
            Path.Join(_tempDirectory, $"ScopedRead_{Guid.NewGuid():N}.xlsx"));

        var setup = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "B2:E4",
            ["values"] = new List<List<object?>>
            {
                new() { "R1B", "R1C", "R1D", "R1E" },
                new() { "R2B", "R2C", "R2D", "R2E" },
                new() { "R3B", "R3C", "R3D", "R3E" }
            }
        });
        AssertSetupSuccess(setup, "range.set-values");

        var response = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "B2:E4",
            ["row_offset"] = 1,
            ["row_limit"] = 1,
            ["columns"] = "E,B"
        });

        AssertSuccess(response, "range.get-values scoped");
        using var document = JsonDocument.Parse(response);
        var root = document.RootElement;
        Assert.Equal(1, root.GetProperty("rowCount").GetInt32());
        Assert.Equal(2, root.GetProperty("columnCount").GetInt32());
        Assert.Equal(3, root.GetProperty("totalRowCount").GetInt32());
        Assert.Equal(4, root.GetProperty("totalColumnCount").GetInt32());
        Assert.Equal(1, root.GetProperty("rowOffset").GetInt32());
        Assert.True(root.GetProperty("hasMoreRows").GetBoolean());
        Assert.Equal(2, root.GetProperty("nextRowOffset").GetInt32());
        Assert.True(root.GetProperty("isTruncated").GetBoolean());
        Assert.Equal(["E", "B"], root.GetProperty("selectedColumns").EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(["R2E", "R2B"], root.GetProperty("values")[0].EnumerateArray().Select(value => value.GetString()));

        await CloseSessionAsync(sessionId, save: false);

        var waitDeadline = DateTime.UtcNow + TimeSpan.FromSeconds(15);
        List<int> leakedExcelProcessIds;
        do
        {
            leakedExcelProcessIds = Process.GetProcessesByName("EXCEL")
                .Select(process => process.Id)
                .Where(processId => !baselineExcelProcessIds.Contains(processId))
                .ToList();

            if (leakedExcelProcessIds.Count == 0)
            {
                break;
            }

            await Task.Delay(TimeSpan.FromMilliseconds(250));
        }
        while (DateTime.UtcNow < waitDeadline);

        Assert.Empty(leakedExcelProcessIds);
    }
}
