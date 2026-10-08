using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Collection("Mac Excel E2E")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "MacHandoff")]
public sealed class MacNativeSessionE2ETests(ITestOutputHelper output)
{
    private static readonly int[][] SavedValues = [[27182]];
    private static readonly int[][] DiscardedValues = [[999]];

    [MacExcelTheory]
    [InlineData("cli", false)]
    [InlineData("cli", true)]
    [InlineData("mcp", false)]
    [InlineData("mcp", true)]
    public async Task NativeOpen_AppliesRequestedVisibilityToTheExactWorkbook(string entryPoint, bool show)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await MacExcelE2ETests.EntryPointClient.CreateAsync(
            MacExcelE2ETests.FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-native-visible-");
        var path = Path.Combine(directory.FullName, $"native-{Guid.NewGuid():N}.xlsx");
        MacExcelE2ETests.CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = MacExcelE2ETests.SessionId(await client.CallAsync(
                "file", "open", null, new() { ["path"] = path, ["show"] = show }, deadline.Token));
            Assert.True(MacNativeWorkbook.IsOpen(path, TimeSpan.FromSeconds(10)));
            using var application = MacAppleEvents.Create(MacAppleEvents.Code("null"), []);
            using var name = MacAppleEvents.Text(Path.GetFileName(path));
            using var workbook = MacAppleEvents.Object(MacExcelDictionary.WorkbookClass, application, MacAppleEvents.Code("name"), name);
            using var first = MacAppleEvents.Create(MacAppleEvents.Code("long"), BitConverter.GetBytes(1));
            using var window = MacAppleEvents.Object(MacExcelDictionary.WindowClass, workbook, MacAppleEvents.Code("indx"), first);
            using var visible = MacAppleEvents.Property(window, MacExcelDictionary.Visible);
            using var appleEvent = MacAppleEvents.Event(MacAppleEvents.Code("core"), MacAppleEvents.Code("getd"));
            MacAppleEvents.Put(appleEvent, MacAppleEvents.Code("----"), visible);
            Assert.Equal(show, MacAppleEvents.Send(appleEvent, TimeSpan.FromSeconds(10))!.GetValue<bool>());

            MacExcelE2ETests.Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            Assert.False(MacNativeWorkbook.IsOpen(path, TimeSpan.FromSeconds(10)));
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
                output.WriteLine($"Failed native visibility run retained its opaque fixture at {directory.FullName}.");
            }
        }
    }

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    public async Task NativeClose_PreservesSaveReopenAndDiscardContracts(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await MacExcelE2ETests.EntryPointClient.CreateAsync(
            MacExcelE2ETests.FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-native-save-");
        var path = Path.Combine(directory.FullName, $"native-{Guid.NewGuid():N}.xlsx");
        MacExcelE2ETests.CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = MacExcelE2ETests.SessionId(await client.CallAsync(
                "file", "open", null, new() { ["path"] = path }, deadline.Token));
            MacExcelE2ETests.Success(await client.CallAsync("range", "set-values", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1", ["values"] = SavedValues }, deadline.Token));
            MacExcelE2ETests.Success(await client.CallAsync("file", "close", session,
                new() { ["save"] = true }, deadline.Token));
            session = null;
            Assert.False(MacNativeWorkbook.IsOpen(path, TimeSpan.FromSeconds(10)));

            session = MacExcelE2ETests.SessionId(await client.CallAsync(
                "file", "open", null, new() { ["path"] = path }, deadline.Token));
            var saved = MacExcelE2ETests.Success(await client.CallAsync("range", "get-values", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1" }, deadline.Token));
            Assert.Equal(27182, saved.GetProperty("values")[0][0].GetDouble());
            MacExcelE2ETests.Success(await client.CallAsync("range", "set-values", session,
                new()
                {
                    ["sheet_name"] = "Sheet1",
                    ["range_address"] = "A1",
                    ["values"] = DiscardedValues,
                    ["overwrite_policy"] = "allow"
                }, deadline.Token));
            MacExcelE2ETests.Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;

            session = MacExcelE2ETests.SessionId(await client.CallAsync(
                "file", "open", null, new() { ["path"] = path }, deadline.Token));
            var discarded = MacExcelE2ETests.Success(await client.CallAsync("range", "get-values", session,
                new() { ["sheet_name"] = "Sheet1", ["range_address"] = "A1" }, deadline.Token));
            Assert.Equal(27182, discarded.GetProperty("values")[0][0].GetDouble());
            MacExcelE2ETests.Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            Assert.False(MacNativeWorkbook.IsOpen(path, TimeSpan.FromSeconds(10)));
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
                output.WriteLine($"Failed native persistence run retained its opaque fixture at {directory.FullName}.");
            }
        }
    }

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    public async Task NativePreflight_RejectsExactPathAndDuplicateNameWithoutOpeningAnotherWorkbook(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await MacExcelE2ETests.EntryPointClient.CreateAsync(
            MacExcelE2ETests.FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-native-preflight-");
        var path = Path.Combine(directory.FullName, $"native-{Guid.NewGuid():N}.xlsx");
        var duplicateDirectory = Directory.CreateDirectory(Path.Combine(directory.FullName, "duplicate"));
        var duplicate = Path.Combine(duplicateDirectory.FullName, Path.GetFileName(path));
        MacExcelE2ETests.CreateBlankWorkbook(path);
        File.Copy(path, duplicate);
        string? session = null;
        var completed = false;
        try
        {
            session = MacExcelE2ETests.SessionId(await client.CallAsync(
                "file", "open", null, new() { ["path"] = path }, deadline.Token));
            var before = MacNativeWorkbook.ReadNames(fullPaths: true, TimeSpan.FromSeconds(10));
            Assert.Contains(MacPathCanonicalizer.Normalize(path), before.Select(MacPathCanonicalizer.Normalize));
            var exact = Assert.Throws<InvalidOperationException>(() =>
                MacNativeWorkbook.PrepareOpen(path, TimeSpan.FromSeconds(10)));
            Assert.Contains("already open", exact.Message, StringComparison.Ordinal);

            var rejected = await client.CallAsync(
                "file", "open", null, new() { ["path"] = duplicate }, deadline.Token);
            Assert.False(rejected.GetProperty("success").GetBoolean());
            Assert.Contains("same name", rejected.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
            var after = MacNativeWorkbook.ReadNames(fullPaths: true, TimeSpan.FromSeconds(10));
            Assert.Equal(before.Order(StringComparer.Ordinal), after.Order(StringComparer.Ordinal));

            MacExcelE2ETests.Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
            var remaining = MacNativeWorkbook.ReadNames(fullPaths: true, TimeSpan.FromSeconds(10));
            Assert.DoesNotContain(MacPathCanonicalizer.Normalize(path), remaining.Select(MacPathCanonicalizer.Normalize));
            completed = true;
        }
        finally
        {
            if (session is not null) await client.TryCloseAsync(session);
            if (completed)
            {
                File.Delete(path);
                File.Delete(duplicate);
                duplicateDirectory.Delete();
                directory.Delete();
            }
            else
            {
                output.WriteLine($"Failed native preflight run retained its opaque fixtures at {directory.FullName}.");
            }
        }
    }
}
