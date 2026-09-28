using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;
using Xunit.Abstractions;
using static Sbroenne.ExcelMcp.Portable.Tests.MacExcelE2ETests;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacNamedRangeTheoryAttribute : TheoryAttribute
{
    public MacNamedRangeTheoryAttribute()
    {
        if (!OperatingSystem.IsMacOS()
            || Environment.GetEnvironmentVariable("EXCELMCP_MAC_E2E") != "1"
            || Environment.GetEnvironmentVariable("EXCELMCP_MAC_NAMED_RANGE_E2E") != "1")
        {
            Skip = "Explicit desktop acceptance required: scripts/Test-MacE2E.ps1 -IncludeNamedRanges.";
        }
    }
}

[Collection("Mac Excel E2E")]
public sealed class MacNamedRangeE2ETests(ITestOutputHelper output)
{
    private static readonly int[][] SentinelValues = [[9876]];
    private static readonly int[][] MatrixValues = [[1, 2], [3, 4]];

    [MacNamedRangeTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    [Trait("Category", "Integration")]
    [Trait("RequiresExcel", "true")]
    [Trait("Feature", "MacNamedRange")]
    public async Task NamedRanges_RealPublicLifecycleAndBoundedPreviews(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3));
        await using var client = await EntryPointClient.CreateAsync(
            FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-mac-names-");
        var path = Path.Combine(directory.FullName, $"names-{Guid.NewGuid():N}.xlsx");
        var sentinelPath = Path.Combine(directory.FullName, $"sentinel-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        CreateBlankWorkbook(sentinelPath);
        AddPreviewFixtures(path);
        string? session = null;
        string? sentinel = null;
        var completed = false;
        const string createdName = "NewInput";
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            sentinel = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = sentinelPath }, deadline.Token));
            Success(await client.CallAsync("range", "set-values", sentinel,
                new() { ["sheet_name"] = "Data", ["range_address"] = "A1", ["values"] = SentinelValues },
                deadline.Token));

            async Task<JsonElement> Call(string action, Dictionary<string, object?> args) =>
                await client.CallAsync("namedrange", action, session, args, deadline.Token);
            var created = Success(await Call("create", new() { ["name"] = createdName, ["reference"] = "Data!$A$1" }));
            Assert.Equal(path, created.GetProperty("filePath").GetString());
            var duplicate = await Call("create", new() { ["name"] = "newinput", ["reference"] = "Data!$B$2" });
            Assert.False(duplicate.GetProperty("success").GetBoolean());
            Assert.Equal("InvalidOperation", duplicate.GetProperty("errorCategory").GetString());
            Success(await Call("write", new() { ["name"] = createdName, ["value"] = "17" }));
            var number = await Call("read", new() { ["name"] = createdName });
            Assert.Equal(17, number.GetProperty("value").GetDouble());
            Assert.Equal("Double", number.GetProperty("valueType").GetString());
            Success(await Call("write", new() { ["name"] = "Data!Input", ["value"] = "99" }));
            Assert.Equal(99, (await Call("read", new() { ["name"] = "Data!Input" })).GetProperty("value").GetDouble());
            Assert.Equal(17, (await Call("read", new() { ["name"] = "Input" })).GetProperty("value").GetDouble());
            Assert.Equal(17, (await Call("read", new() { ["name"] = "DynamicInput" })).GetProperty("value").GetDouble());
            Success(await Call("write", new() { ["name"] = "Data!ShadowOnly", ["value"] = "42" }));
            var collision = await Call("create", new() { ["name"] = "ShadowOnly", ["reference"] = "Data!$A$2" });
            Assert.False(collision.GetProperty("success").GetBoolean());
            Assert.Equal("PlatformNotSupported", collision.GetProperty("errorCategory").GetString());
            Assert.Equal(42, (await Call("read", new() { ["name"] = "Data!ShadowOnly" })).GetProperty("value").GetDouble());
            var scopedCreate = await Call("create", new() { ["name"] = "Data!CreatedLocal", ["reference"] = "Data!$A$2" });
            Assert.False(scopedCreate.GetProperty("success").GetBoolean());
            Assert.Equal("PlatformNotSupported", scopedCreate.GetProperty("errorCategory").GetString());
            var ambiguous = await Call("read", new() { ["name"] = "ShadowedDynamic" });
            Assert.False(ambiguous.GetProperty("success").GetBoolean());
            Assert.Equal("PlatformNotSupported", ambiguous.GetProperty("errorCategory").GetString());
            Success(await client.CallAsync("sheet", "rename", session,
                new() { ["old_name"] = "Spare", ["new_name"] = "Names With Spaces" }, deadline.Token));
            Success(await Call("write", new() { ["name"] = "'Names With Spaces'!LocalInput", ["value"] = "88" }));
            Assert.Equal(88, (await Call("read", new() { ["name"] = "'Names With Spaces'!LocalInput" }))
                .GetProperty("value").GetDouble());
            Success(await client.CallAsync("range", "set-number-format", session,
                new() { ["sheet_name"] = "Data", ["range_address"] = "A1", ["format_code"] = "yyyy-mm-dd" }, deadline.Token));
            Success(await Call("write", new() { ["name"] = createdName, ["value"] = "44927" }));
            var date = await Call("read", new() { ["name"] = createdName });
            Assert.True(date.TryGetProperty("value", out var dateValue), date.GetRawText());
            Assert.Equal(44927, dateValue.GetDouble());
            Success(await client.CallAsync("range", "set-values", session,
                new() { ["sheet_name"] = "", ["range_address"] = createdName, ["values"] = new int[][] { [44928] } },
                deadline.Token));
            foreach (var (action, sheetName, address) in new[]
            {
                ("get-values", "", createdName), ("get-values", "Data", "A1"), ("get-formulas", "Data", "A1")
            })
            {
                var dateRange = Success(await client.CallAsync("range", action, session,
                    new() { ["sheet_name"] = sheetName, ["range_address"] = address }, deadline.Token));
                Assert.Equal(44928, dateRange.GetProperty("values")[0][0].GetDouble());
            }
            Success(await client.CallAsync("range", "set-number-format", session,
                new() { ["sheet_name"] = "Data", ["range_address"] = "A1", ["format_code"] = "General" }, deadline.Token));
            Success(await Call("write", new() { ["name"] = createdName, ["value"] = "true" }));
            var boolean = await Call("read", new() { ["name"] = createdName });
            Assert.True(boolean.GetProperty("value").GetBoolean());
            Assert.Equal("Boolean", boolean.GetProperty("valueType").GetString());
            Success(await Call("write", new() { ["name"] = createdName, ["value"] = "quoted \"text\"" }));
            Assert.Equal("quoted \"text\"", (await Call("read", new() { ["name"] = createdName })).GetProperty("value").GetString());

            Success(await Call("update", new() { ["name"] = createdName, ["reference"] = "==Data!$A$1:$B$2" }));
            Success(await client.CallAsync("range", "set-values", session,
                new() { ["sheet_name"] = "Data", ["range_address"] = "A1:B2", ["values"] = MatrixValues },
                deadline.Token));
            var array = await Call("read", new() { ["name"] = createdName });
            Assert.Equal("Array", array.GetProperty("valueType").GetString());
            Assert.Equal(4, array.GetProperty("value")[1][1].GetDouble());
            var bulk = Success(await client.CallAsync("range", "get-values", session,
                new() { ["sheet_name"] = "", ["range_address"] = createdName }, deadline.Token));
            Assert.Equal(4, bulk.GetProperty("values")[1][1].GetDouble());
            Success(await client.CallAsync("range", "set-values", session,
                new() { ["sheet_name"] = "", ["range_address"] = createdName, ["values"] = MatrixValues },
                deadline.Token));
            Assert.Equal(4, (await Call("read", new() { ["name"] = createdName })).GetProperty("value")[1][1].GetDouble());
            var names = Success(await Call("list", new())).GetProperty("namedRanges").EnumerateArray().ToArray();
            Assert.DoesNotContain(names, item => item.GetProperty("name").GetString() == "HiddenInternal");
            Assert.DoesNotContain(names, item => item.GetProperty("name").GetString()!
                .EndsWith("_FilterDatabase", StringComparison.OrdinalIgnoreCase));
            var boundary = Assert.Single(names, item => item.GetProperty("name").GetString() == "PreviewBoundary");
            Assert.Equal(10_000, boundary.GetProperty("cellCount").GetInt64());
            Assert.Equal("Array", boundary.GetProperty("valueType").GetString());
            Assert.True(boundary.TryGetProperty("value", out _));
            var overBoundary = Assert.Single(names, item => item.GetProperty("name").GetString() == "PreviewOverBoundary");
            Assert.Equal(10_001, overBoundary.GetProperty("cellCount").GetInt64());
            Assert.Equal("RangeTooLarge", overBoundary.GetProperty("valueType").GetString());
            Assert.False(overBoundary.TryGetProperty("value", out _));
            var large = Assert.Single(names, item => item.GetProperty("name").GetString() == "LargePreview");
            Assert.Equal("RangeTooLarge", large.GetProperty("valueType").GetString());
            Assert.Equal(1_048_576, large.GetProperty("cellCount").GetInt64());
            Assert.False(large.TryGetProperty("value", out _));
            var multiple = Assert.Single(names, item => item.GetProperty("name").GetString() == "MultipleAreas");
            Assert.Equal("MultiAreaRange", multiple.GetProperty("valueType").GetString());
            Assert.False(multiple.TryGetProperty("value", out _));
            var constant = Assert.Single(names, item => item.GetProperty("name").GetString() == "ConstantOnly");
            Assert.Equal("Unavailable", constant.GetProperty("valueType").GetString());
            Assert.False(string.IsNullOrWhiteSpace(constant.GetProperty("valueOmittedReason").GetString()));
            var dynamicPreview = Assert.Single(names, item => item.GetProperty("name").GetString() == "ShadowedDynamic");
            Assert.Equal("Unavailable", dynamicPreview.GetProperty("valueType").GetString());
            Assert.Contains("shadowed", dynamicPreview.GetProperty("valueOmittedReason").GetString(), StringComparison.Ordinal);
            Assert.DoesNotContain(names, item => item.GetProperty("name").GetString() == "Data!CreatedLocal");

            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            Assert.Equal(4, (await Call("read", new() { ["name"] = createdName })).GetProperty("value")[1][1].GetDouble());
            Success(await Call("delete", new() { ["name"] = createdName }));
            var missing = await Call("read", new() { ["name"] = createdName });
            Assert.False(missing.GetProperty("success").GetBoolean());
            Assert.Equal("InvalidOperation", missing.GetProperty("errorCategory").GetString());
            Assert.Contains("not found", missing.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
            Success(await Call("update", new() { ["name"] = "'Names With Spaces'!LocalInput", ["reference"] = "Data!$D$1" }));
            Assert.Equal(99, (await Call("read", new() { ["name"] = "'Names With Spaces'!LocalInput" })).GetProperty("value").GetDouble());
            Success(await Call("delete", new() { ["name"] = "'Names With Spaces'!LocalInput" }));
            var missingLocal = await Call("read", new() { ["name"] = "'Names With Spaces'!LocalInput" });
            Assert.Equal("InvalidOperation", missingLocal.GetProperty("errorCategory").GetString());
            Assert.Equal(99, (await Call("read", new() { ["name"] = "Data!Input" })).GetProperty("value").GetDouble());
            var unchanged = Success(await client.CallAsync("range", "get-values", sentinel,
                new() { ["sheet_name"] = "Data", ["range_address"] = "A1" }, deadline.Token));
            Assert.Equal(9876, unchanged.GetProperty("values")[0][0].GetDouble());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            Success(await client.CallAsync("file", "close", sentinel, new(), deadline.Token));
            sentinel = null;
            completed = true;
            output.WriteLine($"{entryPoint}: named-range lifecycle, value types, bounded previews, persistence and sentinel passed.");
        }
        finally
        {
            if (session is not null) await client.TryCloseAsync(session);
            if (sentinel is not null) await client.TryCloseAsync(sentinel);
            if (completed)
            {
                File.Delete(path);
                File.Delete(sentinelPath);
                directory.Delete();
            }
            else
            {
                output.WriteLine($"Failed named-range run retained its synthetic fixtures at {directory.FullName}.");
            }
        }
    }

    private static void AddPreviewFixtures(string path)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = archive.GetEntry("xl/workbook.xml")!;
        XDocument document;
        using (var stream = entry.Open()) document = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        document.Root!.Add(new XElement(ns + "definedNames",
            new XElement(ns + "definedName", new XAttribute("name", "HiddenInternal"), new XAttribute("hidden", "1"), "#REF!"),
            new XElement(ns + "definedName", new XAttribute("name", "_xlnm._FilterDatabase"),
                new XAttribute("localSheetId", "0"), new XAttribute("hidden", "1"), "Data!$A$1"),
            new XElement(ns + "definedName", new XAttribute("name", "PreviewBoundary"), "Data!$A$1:$A$10000"),
            new XElement(ns + "definedName", new XAttribute("name", "PreviewOverBoundary"), "Data!$A$1:$A$10001"),
            new XElement(ns + "definedName", new XAttribute("name", "Input"), new XAttribute("localSheetId", "0"), "Data!$D$1"),
            new XElement(ns + "definedName", new XAttribute("name", "Input"), "Data!$A$1"),
            new XElement(ns + "definedName", new XAttribute("name", "ShadowOnly"), new XAttribute("localSheetId", "0"), "Data!$D$2"),
            new XElement(ns + "definedName", new XAttribute("name", "LocalInput"), new XAttribute("localSheetId", "1"), "Spare!$A$1"),
            new XElement(ns + "definedName", new XAttribute("name", "DynamicInput"), "OFFSET(Data!$A$1,0,0,1,1)"),
            new XElement(ns + "definedName", new XAttribute("name", "ShadowedDynamic"), "OFFSET(Data!$A$1,0,0,1,1)"),
            new XElement(ns + "definedName", new XAttribute("name", "ShadowedDynamic"), new XAttribute("localSheetId", "0"), "Data!$D$3"),
            new XElement(ns + "definedName", new XAttribute("name", "LargePreview"), "Data!$A:$A"),
            new XElement(ns + "definedName", new XAttribute("name", "MultipleAreas"), "Data!$A$1,Data!$B$2"),
            new XElement(ns + "definedName", new XAttribute("name", "ConstantOnly"), "42")));
        entry.Delete();
        using var output = archive.CreateEntry("xl/workbook.xml").Open();
        document.Save(output);
    }
}
