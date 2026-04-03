// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Diagnostics;
using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Regression coverage for worksheet-create followed immediately by non-A1 range writes.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "true")]
public class WorksheetCreateRangeWriteRegressionTests : McpIntegrationTestBase
{
    private readonly string _testExcelFile;
    private string? _sessionId;

    public WorksheetCreateRangeWriteRegressionTests(ITestOutputHelper output)
        : base(output, "WorksheetCreateRangeWriteClient")
    {
        _testExcelFile = Path.Join(CreateTempDirectory("WsCreateWriteRegression"), "WorksheetCreateRangeWrite.xlsx");
    }

    protected override async Task InitializeTestAsync()
    {
        _sessionId = await CreateWorkbookSessionAsync(_testExcelFile);
    }

    [Fact]
    public async Task CreateWorksheet_ThenSetValues_ToNonA1Range_SucceedsViaMcpProtocol()
    {
        var baselineExcelProcessIds = Process.GetProcessesByName("EXCEL")
            .Select(process => process.Id)
            .ToHashSet();

        var sheetName = "Bug2Data";
        var values = new List<List<object?>>
        {
            new() { "R1C1", "R1C2", "R1C3", "R1C4", "R1C5", "R1C6", "R1C7" },
            new() { "R2C1", "R2C2", "R2C3", "R2C4", "R2C5", "R2C6", "R2C7" },
            new() { "R3C1", "R3C2", "R3C3", "R3C4", "R3C5", "R3C6", "R3C7" },
            new() { "R4C1", "R4C2", "R4C3", "R4C4", "R4C5", "R4C6", "R4C7" },
            new() { "R5C1", "R5C2", "R5C3", "R5C4", "R5C5", "R5C6", "R5C7" },
            new() { "R6C1", "R6C2", "R6C3", "R6C4", "R6C5", "R6C6", "R6C7" },
            new() { "R7C1", "R7C2", "R7C3", "R7C4", "R7C5", "R7C6", "R7C7" },
            new() { "R8C1", "R8C2", "R8C3", "R8C4", "R8C5", "R8C6", "R8C7" }
        };

        await CreateWorksheetAsync(_sessionId!, sheetName);

        var setValuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["path"] = _testExcelFile,
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName,
            ["range_address"] = "A3:G10",
            ["values"] = values
        });
        AssertSetupSuccess(setValuesJson, "range.set-values");

        var getValuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["path"] = _testExcelFile,
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName,
            ["range_address"] = "A3:G10"
        });

        using var getValuesDoc = JsonDocument.Parse(getValuesJson);
        var root = getValuesDoc.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean(), $"range.get-values failed: {getValuesJson}");
        Assert.Equal(8, root.GetProperty("rowCount").GetInt32());
        Assert.Equal(7, root.GetProperty("columnCount").GetInt32());

        var returnedValues = root.GetProperty("values");
        for (int rowIndex = 0; rowIndex < values.Count; rowIndex++)
        {
            for (int columnIndex = 0; columnIndex < values[rowIndex].Count; columnIndex++)
            {
                Assert.Equal(values[rowIndex][columnIndex]?.ToString(), returnedValues[rowIndex][columnIndex].GetString());
            }
        }

        var a1Json = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["path"] = _testExcelFile,
            ["session_id"] = _sessionId,
            ["sheet_name"] = sheetName,
            ["range_address"] = "A1"
        });

        using var a1Doc = JsonDocument.Parse(a1Json);
        Assert.True(a1Doc.RootElement.GetProperty("success").GetBoolean(), $"A1 read failed: {a1Json}");
        var a1Cell = a1Doc.RootElement.GetProperty("values")[0][0];
        Assert.True(a1Cell.ValueKind is JsonValueKind.Null || (a1Cell.ValueKind == JsonValueKind.String && string.IsNullOrEmpty(a1Cell.GetString())));

        await CloseSessionAsync(_sessionId, save: false);
        _sessionId = null;

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
