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
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            sentinel = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = sentinelPath }, deadline.Token));
            Success(await client.CallAsync("range", "set-values", sentinel,
                new() { ["sheet_name"] = "Data", ["range_address"] = "A1", ["values"] = SentinelValues },
                deadline.Token));

            async Task<JsonElement> Call(string action, Dictionary<string, object?> args) =>
                await client.CallAsync("namedrange", action, session, args, deadline.Token);
            Success(await Call("create", new() { ["name"] = "Input", ["reference"] = "Data!$A$1" }));
            Assert.False((await Call("create", new() { ["name"] = "input", ["reference"] = "Data!$B$2" }))
                .GetProperty("success").GetBoolean());
            Success(await Call("write", new() { ["name"] = "Input", ["value"] = "17" }));
            var number = await Call("read", new() { ["name"] = "Input" });
            Assert.Equal(17, number.GetProperty("value").GetDouble());
            Assert.Equal("Double", number.GetProperty("valueType").GetString());
            Success(await Call("write", new() { ["name"] = "Input", ["value"] = "true" }));
            var boolean = await Call("read", new() { ["name"] = "Input" });
            Assert.True(boolean.GetProperty("value").GetBoolean());
            Assert.Equal("Boolean", boolean.GetProperty("valueType").GetString());
            Success(await Call("write", new() { ["name"] = "Input", ["value"] = "quoted \"text\"" }));
            Assert.Equal("quoted \"text\"", (await Call("read", new() { ["name"] = "Input" })).GetProperty("value").GetString());

            Success(await Call("update", new() { ["name"] = "Input", ["reference"] = "==Data!$A$1:$B$2" }));
            Success(await client.CallAsync("range", "set-values", session,
                new() { ["sheet_name"] = "Data", ["range_address"] = "A1:B2", ["values"] = MatrixValues },
                deadline.Token));
            var array = await Call("read", new() { ["name"] = "Input" });
            Assert.Equal("Array", array.GetProperty("valueType").GetString());
            Assert.Equal(4, array.GetProperty("value")[1][1].GetDouble());
            var names = Success(await Call("list", new())).GetProperty("namedRanges").EnumerateArray().ToArray();
            Assert.DoesNotContain(names, item => item.GetProperty("name").GetString() == "HiddenInternal");
            Assert.DoesNotContain(names, item => item.GetProperty("name").GetString() == "_FilterDatabase");
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

            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            Assert.Equal(4, (await Call("read", new() { ["name"] = "Input" })).GetProperty("value")[1][1].GetDouble());
            Success(await Call("delete", new() { ["name"] = "Input" }));
            var missing = await Call("read", new() { ["name"] = "Input" });
            Assert.False(missing.GetProperty("success").GetBoolean());
            Assert.Contains("not found", missing.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
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
            new XElement(ns + "definedName", new XAttribute("name", "_FilterDatabase"), "Data!$A$1"),
            new XElement(ns + "definedName", new XAttribute("name", "LargePreview"), "Data!$A:$A"),
            new XElement(ns + "definedName", new XAttribute("name", "MultipleAreas"), "Data!$A$1,Data!$B$2"),
            new XElement(ns + "definedName", new XAttribute("name", "ConstantOnly"), "42")));
        entry.Delete();
        using var output = archive.CreateEntry("xl/workbook.xml").Open();
        document.Save(output);
    }
}
