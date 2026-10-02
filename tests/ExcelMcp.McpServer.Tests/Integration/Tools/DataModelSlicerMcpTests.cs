using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Slicer")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class DataModelSlicerMcpTests(McpProgramTransportFixture fixture) :
    IClassFixture<McpProgramTransportFixture>
{
    private static readonly string[] Captions = ["2026 Q1", "2026 Q2", "2026 Q3"];

    [Fact]
    public async Task DataModelSlicer_RealProtocol_FilterRecoveryAndClearChangePivotValues()
    {
        string session = await fixture.CreateWorkbookSessionAsync(
            fixture.CreateTempPath("DataModelSlicer", ".xlsx"));
        await CallAsync("worksheet", new() { ["action"] = "create", ["sheet_name"] = "SlicerSmoke" });
        await CallAsync("powerquery", new()
        {
            ["action"] = "create",
            ["query_name"] = "SmokeSales",
            ["m_code"] = "#table(type table [Quarter = text, Amount = number], {{\"2026 Q1\", 10}, {\"2026 Q2\", 20}, {\"2026 Q2\", 30}, {\"2026 Q3\", 40}})",
            ["load_destination"] = "data-model"
        });
        await CallAsync("datamodel", new()
        {
            ["action"] = "create-measure",
            ["table_name"] = "SmokeSales",
            ["measure_name"] = "SmokeTotal",
            ["dax_formula"] = "SUM(SmokeSales[Amount])"
        });
        await CallAsync("pivottable", new()
        {
            ["action"] = "create-from-datamodel",
            ["table_name"] = "SmokeSales",
            ["destination_sheet"] = "SlicerSmoke",
            ["destination_cell"] = "A1",
            ["pivot_table_name"] = "SmokePivot"
        });
        foreach (var (action, field) in new[]
        {
            ("add-row-field", "[SmokeSales].[Quarter]"),
            ("add-value-field", "[Measures].[SmokeTotal]")
        })
        {
            await CallAsync("pivottable_field", new()
            {
                ["action"] = action,
                ["pivot_table_name"] = "SmokePivot",
                ["field_name"] = field
            });
        }
        await CallAsync("pivottable", new() { ["action"] = "refresh", ["pivot_table_name"] = "SmokePivot" });
        var created = await CallAsync("slicer", new()
        {
            ["action"] = "create-slicer",
            ["pivot_table_name"] = "SmokePivot",
            ["field_name"] = "[SmokeSales].[Quarter]",
            ["slicer_name"] = "SmokeSlicer",
            ["destination_sheet"] = "SlicerSmoke",
            ["position"] = "D1"
        });
        AssertItems(created, Captions);
        Assert.Equal("SmokePivot", Assert.Single(created.GetProperty("connectedPivotTables")
            .EnumerateArray()).GetString());
        await AssertStateAsync(100, Captions);
        AssertItems(await SelectAsync("[\"2026 Q2\"]"), [Captions[1]]);
        await AssertStateAsync(50, Captions[1]);

        var failed = await CallAsync("slicer", new()
        {
            ["action"] = "set-slicer-selection",
            ["slicer_name"] = "SmokeSlicer",
            ["selected_items"] = "[\"2026 Q1\",\"missing\"]"
        }, success: false);
        Assert.Contains("was not found", failed.GetProperty("errorMessage").GetString());
        await AssertStateAsync(50, Captions[1]);
        AssertItems(await SelectAsync("[\"2026 Q1\"]"), [Captions[0]]);
        await AssertStateAsync(10, Captions[0]);
        AssertItems(await SelectAsync("[\"2026 Q2\"]", false), [Captions[0], Captions[1]]);
        await AssertStateAsync(60, Captions[0], Captions[1]);
        AssertItems(await SelectAsync("[]", false), Captions);
        await AssertStateAsync(100, Captions);
        await fixture.CloseSessionAsync(session);

        async Task<JsonElement> SelectAsync(string selected, bool? clearFirst = null)
        {
            Dictionary<string, object?> arguments = new()
            {
                ["action"] = "set-slicer-selection",
                ["slicer_name"] = "SmokeSlicer",
                ["selected_items"] = selected
            };
            if (clearFirst.HasValue)
                arguments["clear_first"] = clearFirst.Value;
            return await CallAsync("slicer", arguments);
        }

        async Task AssertStateAsync(double total, params string[] selected)
        {
            var listed = await CallAsync("slicer", new()
            {
                ["action"] = "list-slicers",
                ["pivot_table_name"] = "SmokePivot"
            });
            AssertItems(Assert.Single(listed.GetProperty("slicers").EnumerateArray()), selected);
            var values = await CallAsync("range", new()
            {
                ["action"] = "get-values",
                ["sheet_name"] = "SlicerSmoke",
                ["range_address"] = "A1:B6"
            });
            var rows = values.GetProperty("values").EnumerateArray()
                .Where(row => row[1].ValueKind == JsonValueKind.Number).ToArray();
            Assert.Equal(selected.Length + 1, rows.Length);
            Assert.Equal(selected.Order(), rows.Take(selected.Length).Select(row => row[0].GetString()).Order());
            foreach (var row in rows.Take(selected.Length))
            {
                Assert.Equal(row[0].GetString() switch
                {
                    "2026 Q1" => 10d,
                    "2026 Q2" => 50d,
                    "2026 Q3" => 40d,
                    _ => throw new InvalidOperationException("Unexpected PivotTable caption.")
                }, row[1].GetDouble());
            }
            Assert.Equal(total, rows[^1][1].GetDouble());
        }

        async Task<JsonElement> CallAsync(string tool, Dictionary<string, object?> arguments, bool success = true)
        {
            arguments["session_id"] = session;
            string response = await fixture.CallToolAsync(tool, arguments, TimeSpan.FromMinutes(2));
            using var document = JsonDocument.Parse(response);
            Assert.True(document.RootElement.GetProperty("success").GetBoolean() == success, response);
            if (success)
            {
                Assert.True(!document.RootElement.TryGetProperty("errorMessage", out var error)
                    || error.ValueKind == JsonValueKind.Null || error.GetString() == string.Empty, response);
            }
            return document.RootElement.Clone();
        }
    }

    private static void AssertItems(JsonElement slicer, string[] selected)
    {
        Assert.Equal(Captions, slicer.GetProperty("availableItems")
            .EnumerateArray().Select(item => item.GetString()).Order());
        Assert.Equal(selected.Order(), slicer.GetProperty("selectedItems")
            .EnumerateArray().Select(item => item.GetString()).Order());
    }
}
