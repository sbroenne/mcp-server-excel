using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;
using Xunit.Abstractions;
using static Sbroenne.ExcelMcp.Portable.Tests.MacExcelE2ETests;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Collection("Mac Excel E2E")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Worksheets")]
public sealed class MacNativeWorksheetE2ETests(ITestOutputHelper output)
{
    private static readonly string[] CreatedSheetOrder = ["Beta", "Alpha", "Sheet1"];
    private static readonly string[] CreatedSheetNames = ["Alpha", "Beta"];
    private static readonly string[] RejectedNames = ["", " ", "HiStOrY", "'Name", "Name'", "Bad/Name",
        "12345678901234567890123456789012", "sheet1"];

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    public async Task Rename_ReportsExcelRejectionWithSharedDiagnosticAndUnchangedProtectedState(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-protected-name-");
        var path = Path.Combine(directory.FullName, $"protected-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            using (var workbook = MacNativeWorkbook.Resolve(path, TimeSpan.FromSeconds(30)))
            {
                // Fixture-only dictionary API; public workbook protection remains gated.
                using var protect = MacAppleEvents.Event(MacAppleEvents.Code("smXL"), MacAppleEvents.Code("XPTw"));
                MacAppleEvents.Put(protect, MacAppleEvents.Code("----"), workbook);
                using var structure = MacAppleEvents.Create(MacAppleEvents.Code("bool"), [1]);
                MacAppleEvents.Put(protect, MacAppleEvents.Code("5293"), structure);
                MacAppleEvents.SendCommand(protect, TimeSpan.FromSeconds(30));
                Assert.True(MacNativeRange.Read(workbook, MacAppleEvents.Code("1828"), TimeSpan.FromSeconds(30))!.GetValue<bool>());
            }
            var before = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            var rejected = await client.CallAsync("sheet", "rename", session,
                new() { ["old_name"] = "Sheet1", ["new_name"] = "Rejected" }, deadline.Token);
            Assert.False(rejected.GetProperty("success").GetBoolean());
            Assert.Equal("ComInterop", rejected.GetProperty("errorCategory").GetString());
            Assert.Equal("InvalidOperationException", rejected.GetProperty("exceptionType").GetString());
            Assert.Equal("InvalidOperationException: Excel rejected worksheet name 'Rejected'. Worksheet 'Sheet1' was not renamed.",
                rejected.GetProperty("errorMessage").GetString());
            Assert.False(string.IsNullOrEmpty(rejected.GetProperty("innerError").GetString()));
            var unchanged = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal(before.GetProperty("worksheets").GetRawText(), unchanged.GetProperty("worksheets").GetRawText());
            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = null;
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            var reopened = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal(before.GetProperty("worksheets").GetRawText(), reopened.GetProperty("worksheets").GetRawText());
            using (var workbook = MacNativeWorkbook.Resolve(path, TimeSpan.FromSeconds(30)))
                Assert.True(MacNativeRange.Read(workbook, MacAppleEvents.Code("1828"), TimeSpan.FromSeconds(30))!.GetValue<bool>());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
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
            else output.WriteLine($"Failed protected naming acceptance retained its opaque fixture at {directory.FullName}.");
        }
    }

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    public async Task Naming_RejectsChartSheetWorkbooksBeforeMutationAndPreservesSavedState(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-chart-name-gate-");
        var path = Path.Combine(directory.FullName, $"chart-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null,
                new() { ["path"] = path, ["show"] = true }, deadline.Token));
            var chartName = CreateChartSheetFixture(path);
            var before = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal("Sheet1", Assert.Single(before.GetProperty("worksheets").EnumerateArray())
                .GetProperty("name").GetString());
            var create = await client.CallAsync("sheet", "create", session,
                new() { ["sheet_name"] = "Rejected" }, deadline.Token);
            var rename = await client.CallAsync("sheet", "rename", session,
                new() { ["old_name"] = "Sheet1", ["new_name"] = "Rejected" }, deadline.Token);
            foreach (var rejected in new[] { create, rename })
            {
                Assert.False(rejected.GetProperty("success").GetBoolean());
                Assert.Equal("PlatformNotSupported", rejected.GetProperty("errorCategory").GetString());
                Assert.Contains("no naming mutation was attempted", rejected.GetProperty("errorMessage").GetString(),
                    StringComparison.Ordinal);
            }
            var unchanged = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal(before.GetProperty("worksheets").GetRawText(), unchanged.GetProperty("worksheets").GetRawText());
            AssertChartSheet(path, chartName);
            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = null;
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            var reopened = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal(before.GetProperty("worksheets").GetRawText(), reopened.GetProperty("worksheets").GetRawText());
            AssertChartSheet(path, chartName);
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
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
            else output.WriteLine($"Failed chart-sheet gate acceptance retained its opaque fixture at {directory.FullName}.");
        }
    }

    private static string CreateChartSheetFixture(string path)
    {
        using var workbook = MacNativeWorkbook.Resolve(path, TimeSpan.FromSeconds(30));
        using var activeSheet = MacAppleEvents.Property(workbook, MacExcelDictionary.ActiveSheet);
        using var location = MacAppleEvents.Record(MacAppleEvents.Code("insl"));
        using var before = MacAppleEvents.Create(MacAppleEvents.Code("enum"), BitConverter.GetBytes(MacAppleEvents.Code("befo")));
        MacAppleEvents.PutKey(location, MacAppleEvents.Code("kobj"), activeSheet);
        MacAppleEvents.PutKey(location, MacAppleEvents.Code("kpos"), before);
        using var chartClass = MacAppleEvents.Create(MacAppleEvents.Code("type"), BitConverter.GetBytes(MacExcelDictionary.ChartSheetClass));
        using var create = MacAppleEvents.Event(MacExcelDictionary.MakeClass, MacExcelDictionary.MakeId);
        MacAppleEvents.Put(create, MacExcelDictionary.MakeClassParameter, chartClass);
        MacAppleEvents.Put(create, MacExcelDictionary.MakeLocationParameter, location);
        using var chart = MacAppleEvents.SendSpecifier(create, TimeSpan.FromSeconds(30));
        return MacNativeRange.Read(chart, MacExcelDictionary.Name, TimeSpan.FromSeconds(30))
            ?.GetValue<string>() ?? throw new InvalidDataException("Excel did not return the created chart sheet name.");
    }

    private static void AssertChartSheet(string path, string expectedName)
    {
        using var workbook = MacNativeWorkbook.Resolve(path, TimeSpan.FromSeconds(30));
        Assert.Equal(1, MacNativeRange.Count(workbook, MacExcelDictionary.ChartSheetClass, TimeSpan.FromSeconds(30)));
        using var first = MacAppleEvents.Create(MacAppleEvents.Code("long"), BitConverter.GetBytes(1));
        using var chart = MacAppleEvents.Object(MacExcelDictionary.ChartSheetClass, workbook, MacAppleEvents.Code("indx"), first);
        Assert.Equal(expectedName, MacNativeRange.Read(chart, MacExcelDictionary.Name, TimeSpan.FromSeconds(30))?.GetValue<string>());
    }

    [MacExcelTheory]
    [InlineData("cli", false)]
    [InlineData("mcp", false)]
    [InlineData("cli", true)]
    [InlineData("mcp", true)]
    public async Task Naming_RejectsInvalidAndDuplicateNamesBeforeMutationAndPreservesExactSpaces(
        string entryPoint, bool rename)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-name-contract-");
        var path = Path.Combine(directory.FullName, $"names-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            Success(await client.CallAsync("sheet", "create", session, new() { ["sheet_name"] = "Spare" }, deadline.Token));
            var before = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            foreach (var name in RejectedNames)
            {
                var arguments = rename
                    ? new Dictionary<string, object?> { ["old_name"] = "Spare", ["new_name"] = name }
                    : new Dictionary<string, object?> { ["sheet_name"] = name };
                var rejected = await client.CallAsync("sheet", rename ? "rename" : "create", session, arguments, deadline.Token);
                Assert.False(rejected.GetProperty("success").GetBoolean());
                if (string.IsNullOrWhiteSpace(name) && (!rename || entryPoint == "mcp"))
                {
                    Assert.Contains("required", rejected.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
                }
                else
                {
                    Assert.Equal(name == "sheet1" ? "InvalidOperationException" : "ArgumentException",
                        rejected.GetProperty("exceptionType").GetString());
                    Assert.Contains(name == "sheet1" ? "Sheet 'sheet1' already exists." : "Worksheet names must be nonblank",
                        rejected.GetProperty("errorMessage").GetString(), StringComparison.Ordinal);
                }
                var unchanged = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
                Assert.Equal(before.GetProperty("worksheets").GetRawText(), unchanged.GetProperty("worksheets").GetRawText());
            }
            const string spacedName = "  Exact Name  ";
            var validArguments = rename
                ? new Dictionary<string, object?> { ["old_name"] = "Spare", ["new_name"] = spacedName }
                : new Dictionary<string, object?> { ["sheet_name"] = spacedName };
            Success(await client.CallAsync("sheet", rename ? "rename" : "create", session, validArguments, deadline.Token));
            Success(await client.CallAsync("sheet", "rename", session,
                new() { ["old_name"] = spacedName, ["new_name"] = "  EXACT Name  " }, deadline.Token));
            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = null;
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            var reopened = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Contains(reopened.GetProperty("worksheets").EnumerateArray(),
                sheet => sheet.GetProperty("name").GetString() == "  EXACT Name  ");
            Assert.Contains(reopened.GetProperty("worksheets").EnumerateArray(),
                sheet => sheet.GetProperty("name").GetString() == "Sheet1");
            Assert.Equal(rename ? 2 : 3, reopened.GetProperty("worksheets").GetArrayLength());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
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
            else output.WriteLine($"Failed worksheet naming acceptance retained its opaque fixture at {directory.FullName}.");
        }
    }

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    public async Task RenameDelete_RequireExactNamesAndPreserveMissingSheetFailureBeforeMutation(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-exact-sheets-");
        var path = Path.Combine(directory.FullName, $"names-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            Success(await client.CallAsync("sheet", "create", session, new() { ["sheet_name"] = "Spare" }, deadline.Token));
            var before = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            var renamed = await client.CallAsync("sheet", "rename", session,
                new() { ["old_name"] = "sheet1", ["new_name"] = "Changed" }, deadline.Token);
            Assert.False(renamed.GetProperty("success").GetBoolean());
            Assert.Equal("InvalidOperationException: Sheet 'sheet1' not found.", renamed.GetProperty("errorMessage").GetString());
            Assert.Equal("InvalidOperationException", renamed.GetProperty("exceptionType").GetString());
            var deleted = await client.CallAsync("sheet", "delete", session,
                new() { ["sheet_name"] = "'Sheet1'" }, deadline.Token);
            Assert.False(deleted.GetProperty("success").GetBoolean());
            Assert.Equal("InvalidOperationException: Sheet ''Sheet1'' not found.", deleted.GetProperty("errorMessage").GetString());
            Assert.Equal("InvalidOperationException", deleted.GetProperty("exceptionType").GetString());
            var unchanged = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal(before.GetProperty("worksheets").GetRawText(), unchanged.GetProperty("worksheets").GetRawText());
            var validRename = Success(await client.CallAsync("sheet", "rename", session,
                new() { ["old_name"] = "Spare", ["new_name"] = "Renamed" }, deadline.Token));
            Assert.Equal(path, validRename.GetProperty("filePath").GetString());
            var validDelete = Success(await client.CallAsync("sheet", "delete", session,
                new() { ["sheet_name"] = "Renamed" }, deadline.Token));
            Assert.Equal(path, validDelete.GetProperty("filePath").GetString());
            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = null;
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            var reopened = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal("Sheet1", Assert.Single(reopened.GetProperty("worksheets").EnumerateArray()).GetProperty("name").GetString());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
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
                output.WriteLine($"Failed exact-name acceptance retained its opaque fixture at {directory.FullName}.");
            }
        }
    }

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    public async Task Create_InsertsBeforeTheActiveSheetAndPersistsTheWindowsOrder(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-native-create-");
        var path = Path.Combine(directory.FullName, $"create-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            foreach (var name in CreatedSheetNames)
            {
                var created = Success(await client.CallAsync("sheet", "create", session, new() { ["sheet_name"] = name }, deadline.Token));
                Assert.Equal(path, created.GetProperty("filePath").GetString());
            }
            var listed = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal(CreatedSheetOrder,
                listed.GetProperty("worksheets").EnumerateArray().Select(sheet => sheet.GetProperty("name").GetString()));
            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = null;
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            var reopened = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal(listed.GetProperty("worksheets").GetRawText(), reopened.GetProperty("worksheets").GetRawText());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
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
                output.WriteLine($"Failed native create run retained its opaque fixture at {directory.FullName}.");
            }
        }
    }

    [MacExcelTheory]
    [InlineData("cli")]
    [InlineData("mcp")]
    public async Task NativeList_ReturnsNamesIndicesAndAllVisibilityStatesFromTheOwnedWorkbook(string entryPoint)
    {
        Assert.Equal(0, MacAutomationAccess.Check());
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
        await using var client = await EntryPointClient.CreateAsync(FindRepository(), entryPoint, output, deadline.Token);
        var directory = Directory.CreateTempSubdirectory("excelmcp-native-sheets-");
        var path = Path.Combine(directory.FullName, $"sheets-{Guid.NewGuid():N}.xlsx");
        CreateBlankWorkbook(path);
        string? session = null;
        var completed = false;
        try
        {
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            Success(await client.CallAsync("sheet", "create", session, new() { ["sheet_name"] = "Hidden" }, deadline.Token));
            Success(await client.CallAsync("sheet", "create", session, new() { ["sheet_name"] = "VeryHidden" }, deadline.Token));
            Success(await client.CallAsync("worksheetstyle", "hide", session, new() { ["sheet_name"] = "Hidden" }, deadline.Token));
            Success(await client.CallAsync("worksheetstyle", "very-hide", session, new() { ["sheet_name"] = "VeryHidden" }, deadline.Token));
            var listed = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal(path, listed.GetProperty("filePath").GetString());
            var worksheets = listed.GetProperty("worksheets").EnumerateArray().ToArray();
            Assert.Equal(3, worksheets.Length);
            Assert.Equal([1, 2, 3], worksheets.Select(sheet => sheet.GetProperty("index").GetInt32()));
            Assert.True(worksheets.Single(sheet => sheet.GetProperty("name").GetString() == "Sheet1").GetProperty("visible").GetBoolean());
            Assert.False(worksheets.Single(sheet => sheet.GetProperty("name").GetString() == "Hidden").GetProperty("visible").GetBoolean());
            Assert.False(worksheets.Single(sheet => sheet.GetProperty("name").GetString() == "VeryHidden").GetProperty("visible").GetBoolean());
            Success(await client.CallAsync("file", "close", session, new() { ["save"] = true }, deadline.Token));
            session = null;
            session = SessionId(await client.CallAsync("file", "open", null, new() { ["path"] = path }, deadline.Token));
            var reopened = Success(await client.CallAsync("sheet", "list", session, new(), deadline.Token));
            Assert.Equal(listed.GetProperty("worksheets").GetRawText(), reopened.GetProperty("worksheets").GetRawText());
            Success(await client.CallAsync("file", "close", session, new(), deadline.Token));
            session = null;
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
                output.WriteLine($"Failed native worksheet run retained its opaque fixture at {directory.FullName}.");
            }
        }
    }
}
