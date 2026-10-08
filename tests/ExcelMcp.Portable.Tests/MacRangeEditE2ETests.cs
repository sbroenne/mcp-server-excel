using System.Text.Json;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;
using Xunit.Abstractions;
using static Sbroenne.ExcelMcp.Portable.Tests.MacExcelE2ETests;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Collection("Mac Excel E2E")]
public sealed class MacRangeEditE2ETests(ITestOutputHelper output)
{
    [MacExcelTheory]
    [InlineData("cli", "insert-cells", "Down")]
    [InlineData("cli", "insert-cells", "Right")]
    [InlineData("cli", "delete-cells", "Up")]
    [InlineData("cli", "delete-cells", "Left")]
    [InlineData("cli", "insert-rows", null)]
    [InlineData("cli", "delete-rows", null)]
    [InlineData("cli", "insert-columns", null)]
    [InlineData("cli", "delete-columns", null)]
    [InlineData("mcp", "insert-cells", "Down")]
    [InlineData("mcp", "insert-cells", "Right")]
    [InlineData("mcp", "delete-cells", "Up")]
    [InlineData("mcp", "delete-cells", "Left")]
    [InlineData("mcp", "insert-rows", null)]
    [InlineData("mcp", "delete-rows", null)]
    [InlineData("mcp", "insert-columns", null)]
    [InlineData("mcp", "delete-columns", null)]
    [Trait("Category", "Integration")]
    [Trait("RequiresExcel", "true")]
    [Trait("Feature", "MacRangeEdit")]
    public async Task StructuralEdit_PreservesValuesFormulasAndOwnedWorkbook(
        string entryPoint,
        string action,
        string? shift)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(
            FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-mac-edit-");
        var path = Path.Combine(directory.FullName, $"edit-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync(
                "file", "open", null, new() { ["path"] = path }, deadline.Token));
            await InitializeWorkbookAsync(client, session, includeSpare: true, deadline.Token);

            async Task<JsonElement> Read(string address, string sheet = "Data") =>
                Success(await client.CallAsync("range", "get-values", session,
                    new() { ["sheet_name"] = sheet, ["range_address"] = address }, deadline.Token));

            Success(await client.CallAsync("range", "set-values", session,
                new()
                {
                    ["sheet_name"] = "Data",
                    ["range_address"] = "A1:D4",
                    ["values"] = new int[][] { [11, 12, 13, 14], [21, 22, 23, 24], [31, 32, 33, 34], [41, 42, 43, 44] }
                }, deadline.Token));
            Success(await client.CallAsync("range", "set-formulas", session,
                new() { ["sheet_name"] = "Data", ["range_address"] = "H8", ["formulas"] = new string[][] { ["=SUM(A1:D4)"] } },
                deadline.Token));
            Success(await client.CallAsync("range", "set-values", session,
                new() { ["sheet_name"] = "Spare", ["range_address"] = "A1", ["values"] = new int[][] { [9876] } },
                deadline.Token));

            var address = action.EndsWith("rows", StringComparison.Ordinal) ? "2:3"
                : action.EndsWith("columns", StringComparison.Ordinal) ? "B:C" : "B2:C3";
            var arguments = new Dictionary<string, object?>
            {
                ["sheet_name"] = "Data",
                ["range_address"] = address
            };
            if (shift is not null)
            {
                arguments[action == "insert-cells" ? "insert_shift" : "delete_shift"] = shift;
            }
            var edited = Success(await client.CallAsync("rangeedit", action, session, arguments, deadline.Token));
            Assert.Equal(path, edited.GetProperty("filePath").GetString());
            Assert.Equal(action, edited.GetProperty("action").GetString());
            var vertical = shift is "Down" or "Up" || action.EndsWith("rows", StringComparison.Ordinal);
            var inserted = action.StartsWith("insert", StringComparison.Ordinal);
            var moved = await Read(vertical ? inserted ? "B4:C5" : "B2:C2" : inserted ? "D2:E3" : "B2:B3");
            var expectedValues = inserted ? new int[][] { [22, 23], [32, 33] }
                : vertical ? new int[][] { [42, 43] } : new int[][] { [24], [34] };
            for (var row = 0; row < expectedValues.Length; row++)
            {
                Assert.Equal(expectedValues[row].Select(value => (double)value),
                    moved.GetProperty("values")[row].EnumerateArray().Select(value => value.GetDouble()));
            }
            if (inserted)
            {
                var blank = await Read("B2:C3");
                Assert.All(blank.GetProperty("values").EnumerateArray(),
                    row => Assert.All(row.EnumerateArray(), cell => Assert.Equal(string.Empty, cell.GetString())));
            }
            Assert.Equal(11, (await Read("A1")).GetProperty("values")[0][0].GetDouble());
            Assert.Equal(9876, (await Read("A1", "Spare")).GetProperty("values")[0][0].GetDouble());
            var formulaAddress = action.EndsWith("rows", StringComparison.Ordinal) ? inserted ? "H10" : "H6"
                : action.EndsWith("columns", StringComparison.Ordinal) ? inserted ? "J8" : "F8" : "H8";
            var formulas = Success(await client.CallAsync("range", "get-formulas", session,
                new() { ["sheet_name"] = "Data", ["range_address"] = formulaAddress }, deadline.Token));
            Assert.StartsWith("=SUM(", formulas.GetProperty("formulas")[0][0].GetString(), StringComparison.Ordinal);
            if (action.EndsWith("rows", StringComparison.Ordinal))
            {
                Assert.Equal(inserted ? "=SUM(A1:D6)" : "=SUM(A1:D2)",
                    formulas.GetProperty("formulas")[0][0].GetString());
            }
            if (action.EndsWith("columns", StringComparison.Ordinal))
            {
                Assert.Equal(inserted ? "=SUM(A1:F4)" : "=SUM(A1:B4)",
                    formulas.GetProperty("formulas")[0][0].GetString());
            }

            var beforeInvalid = (await Read("A1:D4")).GetProperty("values").GetRawText();
            var invalid = new Dictionary<string, object?>(arguments) { ["range_address"] = "not-a-range" };
            var rejected = await client.CallAsync("rangeedit", action, session, invalid, deadline.Token);
            Assert.False(rejected.GetProperty("success").GetBoolean());
            Assert.False(string.IsNullOrWhiteSpace(rejected.GetProperty("errorMessage").GetString()));
            Assert.Equal(beforeInvalid, (await Read("A1:D4")).GetProperty("values").GetRawText());
            if (shift is null)
            {
                var disjoint = new Dictionary<string, object?>(arguments)
                {
                    ["range_address"] = vertical ? "2:2,4:4" : "B:B,D:D"
                };
                var unsupported = await client.CallAsync("rangeedit", action, session, disjoint, deadline.Token);
                Assert.False(unsupported.GetProperty("success").GetBoolean());
                Assert.Equal("PlatformNotSupported", unsupported.GetProperty("errorCategory").GetString());
                Assert.Equal(beforeInvalid, (await Read("A1:D4")).GetProperty("values").GetRawText());
            }

            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            Assert.Equal(moved.GetProperty("values").GetRawText(),
                (await Read(vertical ? inserted ? "B4:C5" : "B2:C2" : inserted ? "D2:E3" : "B2:B3"))
                    .GetProperty("values").GetRawText());
            Assert.Equal(9876, (await Read("A1", "Spare")).GetProperty("values")[0][0].GetDouble());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            var inventory = Success(await client.CallAsync("file", "list", null, new(), deadline.Token));
            Assert.Empty(inventory.GetProperty("sessions").EnumerateArray());
            completed = true;
        }
        finally
        {
            if (session is not null) await client.TryCloseAsync(session);
            if (completed)
            {
                File.Delete(path);
                directory.Delete();
            }
            else
            {
                output.WriteLine($"Failed structural-edit acceptance retained its opaque workbook at {path}.");
            }
        }
    }
}
