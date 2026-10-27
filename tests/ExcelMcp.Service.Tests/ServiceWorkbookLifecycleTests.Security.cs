using System.Text.Json;
using Microsoft.Win32;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Vbe = Microsoft.Vbe.Interop;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task SecurityNotices_InspectWithoutChangingUnsavedWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var id = await CreateSessionAsync(service, Path.Join(directory, "notices.xlsx"));
            sessions[id] = 0;
            await WriteMarkerAsync(service, id, "keep unsaved");
            var response = await SendSecurityAsync(service, id, "notices");
            RequireSuccess(response);
            using var result = JsonDocument.Parse(response.Result!);
            var notices = result.RootElement.GetProperty("notices");
            Assert.False(notices.GetProperty("protectedView").GetBoolean());
            Assert.False(notices.GetProperty("readOnly").GetBoolean());
            Assert.False(notices.GetProperty("hasVbaProject").GetBoolean());
            Assert.Equal(0, notices.GetProperty("connectionCount").GetInt32());
            Assert.Equal("unknown", notices.GetProperty("vbaState").GetString());
            Assert.Equal("unknown", notices.GetProperty("connectionsState").GetString());
            Assert.False(notices.GetProperty("complete").GetBoolean());
            Assert.Contains(notices.GetProperty("messages").EnumerateArray(),
                message => message.GetString()!.Contains("not expose every notification", StringComparison.Ordinal));
            Assert.Equal("keep unsaved", await ReadMarkerAsync(service, id));
            Assert.False(service.SessionManager.GetSession(id)!.Execute((ctx, _) => ctx.Book.Saved));
        });
    }

    [Theory]
    [InlineData("enable-editing", false, false)]
    [InlineData("enable-editing", true, false)]
    [InlineData("enable-connections", false, false)]
    [InlineData("enable-connections", true, false)]
    [InlineData("enable-macros", false, false)]
    [InlineData("enable-macros", true, false)]
    public async Task SecurityAction_MissingApprovalPreservesWorkbook(string action, bool approved, bool extraApproval)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var id = await CreateSessionAsync(service, Path.Join(directory, "approval.xlsx"));
            sessions[id] = 0;
            await WriteMarkerAsync(service, id, "do not change");
            var args = action == "enable-macros"
                ? (object)new { approved, allowReopen = extraApproval }
                : new { approved, allowRememberedTrust = extraApproval };
            var response = await SendSecurityAsync(service, id, action, args);
            Assert.False(response.Success);
            Assert.Contains("approval", response.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
            Assert.Equal("do not change", await ReadMarkerAsync(service, id));
            Assert.False(service.SessionManager.GetSession(id)!.Execute((ctx, _) => ctx.Book.Saved));
        });
    }

    [Fact]
    public async Task SecurityMacros_DisabledOpeningAndExplicitSavedReopenPreserveSession()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "startup-code.xlsm");
            var created = await CreateSessionAsync(service, path);
            sessions[created] = 0;
            await WriteMarkerAsync(service, created, "not-run");
            service.SessionManager.GetSession(created)!.Execute((ctx, _) => InstallSecurityProbeMacros(ctx));
            await CloseSessionAsync(service, created, save: true);
            sessions.TryRemove(created, out _);

            var id = await OpenSessionAsync(service, path);
            sessions[id] = 0;
            Assert.Equal("not-run", await ReadMarkerAsync(service, id));
            var blockedRun = await service.ProcessAsync(new ServiceRequest
            {
                Command = "vba.run",
                SessionId = id,
                Args = """{"procedureName":"ThisWorkbook.SecurityMarker"}"""
            });
            Assert.False(blockedRun.Success);
            Assert.Contains("allow_reopen", blockedRun.ErrorMessage!, StringComparison.Ordinal);
            Assert.Equal("not-run", await ReadMarkerAsync(service, id));
            var before = service.SessionManager.GetSession(id)!;
            var processId = before.ExcelProcessId;
            var settings = before.Execute((ctx, _) => (ctx.App.AutomationSecurity, ctx.App.DisplayAlerts));
            var inspection = await SendSecurityAsync(service, id, "notices");
            RequireSuccess(inspection);
            using (var json = JsonDocument.Parse(inspection.Result!))
            {
                var notices = json.RootElement.GetProperty("notices");
                Assert.True(notices.GetProperty("hasVbaProject").GetBoolean());
                Assert.Equal("disabledAtOpen", notices.GetProperty("vbaState").GetString());
            }
            await WriteMarkerAsync(service, id, "unsaved");
            var refused = await SendSecurityAsync(service, id, "enable-macros", new { approved = true, allowReopen = true });
            Assert.False(refused.Success);
            Assert.Contains("unsaved", refused.ErrorMessage!, StringComparison.Ordinal);
            Assert.Equal("unsaved", await ReadMarkerAsync(service, id));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "session.save", SessionId = id }));

            var enabled = await SendSecurityAsync(service, id, "enable-macros", new { approved = true, allowReopen = true });
            RequireSuccess(enabled);
            using (var json = JsonDocument.Parse(enabled.Result!))
            {
                var security = json.RootElement.GetProperty("security");
                Assert.True(security.GetProperty("changed").GetBoolean());
                Assert.True(security.GetProperty("reopened").GetBoolean());
                Assert.False(security.GetProperty("trustMayBeRemembered").GetBoolean());
            }
            Assert.Equal("opening-ran", await ReadMarkerAsync(service, id));
            Assert.Same(before, service.SessionManager.GetSession(id));
            Assert.Equal(processId, before.ExcelProcessId);
            Assert.Equal(settings, before.Execute((ctx, _) => (ctx.App.AutomationSecurity, ctx.App.DisplayAlerts)));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "vba.run",
                SessionId = id,
                Args = """{"procedureName":"ThisWorkbook.SecurityMarker"}"""
            }));
            Assert.Equal("explicit-ran-3", await ReadMarkerAsync(service, id));
            var repeated = await SendSecurityAsync(service, id, "enable-macros", new { approved = true, allowReopen = true });
            RequireSuccess(repeated);
            using (var json = JsonDocument.Parse(repeated.Result!))
                Assert.False(json.RootElement.GetProperty("security").GetProperty("reopened").GetBoolean());
            Assert.Equal("explicit-ran-3", await ReadMarkerAsync(service, id));
        });
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task SecurityEditing_ProtectedViewDiscoveryAndApprovedTransitionPreserveNeighbor(bool enableEditing, bool hasMacros)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var protectedPath = Path.Join(directory, hasMacros ? "protected-view.xlsm" : "protected-view.xlsx");
            if (hasMacros)
            {
                var created = await CreateSessionAsync(service, protectedPath);
                sessions[created] = 0;
                await WriteMarkerAsync(service, created, "not-run");
                service.SessionManager.GetSession(created)!.Execute((ctx, _) => InstallSecurityProbeMacros(ctx));
                await CloseSessionAsync(service, created, save: true);
                sessions.TryRemove(created, out _);
            }
            else
                SavedWorkbookTemplates.CopyBlankTo(protectedPath);
            const string trustKey = @"Software\Microsoft\Office\16.0\Excel\Security\Trusted Documents\TrustRecords";
            using (var before = Registry.CurrentUser.OpenSubKey(trustKey))
                Assert.Empty(SyntheticTrustNames(before, protectedPath));
            try
            {
                var id = await CreateSessionAsync(service, Path.Join(directory, "neighbor.xlsx"));
                sessions[id] = 0;
                await WriteMarkerAsync(service, id, "preserve neighbor");
                var batch = service.SessionManager.GetSession(id)!;
                batch.Execute((ctx, _) =>
                {
                    Excel.ProtectedViewWindows? windows = null;
                    Excel.ProtectedViewWindow? view = null;
                    try
                    {
                        windows = ctx.App.ProtectedViewWindows;
                        view = windows.Open(protectedPath);
                    }
                    finally
                    {
                        ComUtilities.Release(ref view);
                        ComUtilities.Release(ref windows);
                    }
                });
                var discovery = await service.ProcessAsync(new ServiceRequest { Command = "session.discover" });
                RequireSuccess(discovery);
                using var discoveryJson = JsonDocument.Parse(discovery.Result!);
                Assert.True(discoveryJson.RootElement.GetProperty("complete").GetBoolean(), discovery.Result);
                var selection = Assert.Single(discoveryJson.RootElement.GetProperty("workbooks").EnumerateArray(),
                    item => item.GetProperty("name").GetString() == Path.GetFileName(protectedPath)).Clone();
                Assert.True(selection.GetProperty("protectedView").GetBoolean());
                var attachedId = await AttachReviewWorkbookAsync(service, selection);
                sessions[attachedId] = 0;
                var listing = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
                RequireSuccess(listing);
                using (var json = JsonDocument.Parse(listing.Result!))
                    Assert.True(Assert.Single(json.RootElement.GetProperty("sessions").EnumerateArray(),
                        item => item.GetProperty("sessionId").GetString() == attachedId).GetProperty("canClose").GetBoolean());
                var status = await SendSecurityAsync(service, attachedId, "notices");
                RequireSuccess(status);
                using (var json = JsonDocument.Parse(status.Result!))
                {
                    var notices = json.RootElement.GetProperty("notices");
                    Assert.True(notices.GetProperty("protectedView").GetBoolean());
                    Assert.Equal(JsonValueKind.Null, notices.GetProperty("hasVbaProject").ValueKind);
                    Assert.Equal(JsonValueKind.Null, notices.GetProperty("connectionCount").ValueKind);
                }
                var attachedMacros = await SendSecurityAsync(service, attachedId, "enable-macros",
                    new { approved = true, allowReopen = true });
                Assert.False(attachedMacros.Success);
                Assert.Contains("attached", attachedMacros.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
                var denied = await SendSecurityAsync(service, attachedId, "enable-editing", new { approved = true });
                Assert.False(denied.Success);
                Assert.Contains("remember", denied.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
                if (!enableEditing)
                {
                    await CloseSessionAsync(service, id, save: false);
                    sessions.TryRemove(id, out _);
                    var surviving = await SendSecurityAsync(service, attachedId, "notices");
                    RequireSuccess(surviving);
                    using (var json = JsonDocument.Parse(surviving.Result!))
                        Assert.True(json.RootElement.GetProperty("notices").GetProperty("protectedView").GetBoolean());
                    await CloseSessionAsync(service, attachedId, save: false);
                    sessions.TryRemove(attachedId, out _);
                    return;
                }
                var originalSecurity = batch.Execute((ctx, _) => ctx.App.AutomationSecurity);
                try
                {
                    batch.Execute((ctx, _) => ctx.App.AutomationSecurity = Office.MsoAutomationSecurity.msoAutomationSecurityLow);
                    var enabled = await SendSecurityAsync(service, attachedId, "enable-editing",
                        new { approved = true, allowRememberedTrust = true });
                    RequireSuccess(enabled);
                    using (var json = JsonDocument.Parse(enabled.Result!))
                    {
                        var security = json.RootElement.GetProperty("security");
                        Assert.True(security.GetProperty("changed").GetBoolean());
                        Assert.True(security.GetProperty("trustMayBeRemembered").GetBoolean());
                        Assert.False(security.GetProperty("reopened").GetBoolean());
                    }
                    var after = await SendSecurityAsync(service, attachedId, "notices");
                    RequireSuccess(after);
                    using (var json = JsonDocument.Parse(after.Result!))
                    {
                        var notices = json.RootElement.GetProperty("notices");
                        Assert.False(notices.GetProperty("protectedView").GetBoolean());
                        Assert.False(notices.GetProperty("readOnly").GetBoolean());
                        Assert.Equal("disabledAtOpen", notices.GetProperty("vbaState").GetString());
                        if (hasMacros)
                        {
                            var messages = notices.GetProperty("messages").EnumerateArray()
                                .Select(message => message.GetString()!).ToArray();
                            Assert.Contains(messages, message => message.Contains("Use Excel", StringComparison.Ordinal));
                            Assert.DoesNotContain(messages, message => message.Contains("requires a separately approved reopen", StringComparison.Ordinal));
                        }
                    }
                    if (hasMacros)
                    {
                        Assert.Equal("not-run", await ReadMarkerAsync(service, attachedId));
                        var blockedRun = await service.ProcessAsync(new ServiceRequest
                        {
                            Command = "vba.run",
                            SessionId = attachedId,
                            Args = """{"procedureName":"ThisWorkbook.SecurityMarker"}"""
                        });
                        Assert.False(blockedRun.Success);
                        Assert.Contains("Use Excel", blockedRun.ErrorMessage!, StringComparison.Ordinal);
                        Assert.Contains("already open", blockedRun.ErrorMessage!, StringComparison.Ordinal);
                        Assert.DoesNotContain("allow_reopen", blockedRun.ErrorMessage!, StringComparison.Ordinal);
                        Assert.Equal("not-run", await ReadMarkerAsync(service, attachedId));
                        await WriteMarkerAsync(service, attachedId, "editing allowed");
                        Assert.Equal("editing allowed", await ReadMarkerAsync(service, attachedId));
                    }
                    Assert.Equal("preserve neighbor", await ReadMarkerAsync(service, id));
                    Assert.False(batch.Execute((ctx, _) => ctx.Book.Saved));
                    Assert.Equal(Office.MsoAutomationSecurity.msoAutomationSecurityLow,
                        batch.Execute((ctx, _) => ctx.App.AutomationSecurity));
                }
                finally
                {
                    batch.Execute((ctx, _) => ctx.App.AutomationSecurity = originalSecurity);
                }
                await CloseSessionAsync(service, attachedId, save: false);
                sessions.TryRemove(attachedId, out _);
            }
            finally
            {
                using var trusted = Registry.CurrentUser.OpenSubKey(trustKey, writable: true);
                foreach (var name in SyntheticTrustNames(trusted, protectedPath))
                    trusted!.DeleteValue(name, throwOnMissingValue: true);
            }
        });
    }

    private static string[] SyntheticTrustNames(RegistryKey? key, string path) =>
        key?.GetValueNames().Where(name => string.Equals(
            Environment.ExpandEnvironmentVariables(name).Replace('/', '\\'), path,
            StringComparison.OrdinalIgnoreCase)).ToArray() ?? [];

    [Fact]
    public async Task SecurityConnections_NoConnectionsReportsNoChangeAndEditingDoesNotRemoveOtherRestrictions()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var id = await CreateSessionAsync(service, Path.Join(directory, "connections.xlsx"));
            sessions[id] = 0;
            var enabled = await SendSecurityAsync(service, id, "enable-connections",
                new { approved = true, allowRememberedTrust = true });
            RequireSuccess(enabled);
            using (var json = JsonDocument.Parse(enabled.Result!))
            {
                var security = json.RootElement.GetProperty("security");
                Assert.False(security.GetProperty("changed").GetBoolean());
                Assert.False(security.GetProperty("trustMayBeRemembered").GetBoolean());
            }
            var editing = await SendSecurityAsync(service, id, "enable-editing",
                new { approved = true, allowRememberedTrust = true });
            Assert.False(editing.Success);
            Assert.Contains("not in Protected View", editing.ErrorMessage!, StringComparison.Ordinal);
            Assert.True(service.SessionManager.GetSession(id)!.Execute((ctx, _) => ctx.Book.Saved));
        });
    }

    [Fact]
    public async Task SecurityConnections_NonemptyWorkbookReportsRequestWithoutClaimingVerifiedEnablement()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "connection-request.xlsx");
            var id = await CreateSessionAsync(service, path);
            sessions[id] = 0;
            await WriteMarkerAsync(service, id, "preserve workbook");
            RequireSuccess(await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.create",
                SessionId = id,
                Args = """{"queryName":"NoticeData","mCode":"#table({\"Marker\"}, {{\"connection-data\"}})","loadDestination":"worksheet","targetSheet":"NoticeData"}"""
            }));
            var before = await SendSecurityAsync(service, id, "notices");
            RequireSuccess(before);
            using (var json = JsonDocument.Parse(before.Result!))
            {
                var notices = json.RootElement.GetProperty("notices");
                Assert.Equal(1, notices.GetProperty("connectionCount").GetInt32());
                Assert.Equal("unknown", notices.GetProperty("connectionsState").GetString());
            }
            ServiceResponse enabled;
            try
            {
                enabled = await SendSecurityAsync(service, id, "enable-connections",
                    new { approved = true, allowRememberedTrust = true });
            }
            finally
            {
                using var trusted = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Office\16.0\Excel\Security\Trusted Documents\TrustRecords", writable: true);
                foreach (var name in SyntheticTrustNames(trusted, path))
                    trusted!.DeleteValue(name, throwOnMissingValue: true);
            }
            RequireSuccess(enabled);
            using (var json = JsonDocument.Parse(enabled.Result!))
            {
                var security = json.RootElement.GetProperty("security");
                Assert.Equal("enable-connections", security.GetProperty("action").GetString());
                Assert.Equal(JsonValueKind.Null, security.GetProperty("changed").ValueKind);
                Assert.True(security.GetProperty("trustMayBeRemembered").GetBoolean());
                Assert.False(security.GetProperty("reopened").GetBoolean());
                Assert.Contains("completion was not established", security.GetProperty("message").GetString()!, StringComparison.Ordinal);
            }
            var after = await SendSecurityAsync(service, id, "notices");
            RequireSuccess(after);
            using (var json = JsonDocument.Parse(after.Result!))
            {
                var notices = json.RootElement.GetProperty("notices");
                Assert.Equal(1, notices.GetProperty("connectionCount").GetInt32());
                Assert.Equal("enableRequested", notices.GetProperty("connectionsState").GetString());
                Assert.False(notices.GetProperty("complete").GetBoolean());
            }
            var values = await service.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = id,
                Args = """{"sheetName":"NoticeData","rangeAddress":"A1:A2"}"""
            });
            RequireSuccess(values);
            using (var json = JsonDocument.Parse(values.Result!))
            {
                var rows = json.RootElement.GetProperty("values");
                Assert.Equal("Marker", rows[0][0].GetString());
                Assert.Equal("connection-data", rows[1][0].GetString());
            }
            Assert.Equal("preserve workbook", await ReadMarkerAsync(service, id));
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SecurityMacros_NeighborPreventsReopenWithoutChangingEitherWorkbook(bool protectedView)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var id = await CreateSessionAsync(service, Path.Join(directory, "macro-owner.xlsm"));
            sessions[id] = 0;
            await WriteMarkerAsync(service, id, "saved original");
            var owner = service.SessionManager.GetSession(id)!;
            owner.Execute((ctx, _) => InstallSecurityProbeMacros(ctx));
            RequireSuccess(await service.ProcessAsync(new ServiceRequest { Command = "session.save", SessionId = id }));
            var neighborPath = Path.Join(directory, "macro-neighbor.xlsx");
            SavedWorkbookTemplates.CopyBlankTo(neighborPath);
            owner.Execute((ctx, _) =>
            {
                Excel.Workbooks? books = null;
                Excel.Workbook? neighbor = null;
                Excel.ProtectedViewWindows? windows = null;
                Excel.ProtectedViewWindow? view = null;
                try
                {
                    if (protectedView)
                    {
                        windows = ctx.App.ProtectedViewWindows;
                        view = windows.Open(neighborPath);
                    }
                    else
                    {
                        books = ctx.App.Workbooks;
                        neighbor = books.Open(neighborPath, UpdateLinks: 0);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref view);
                    ComUtilities.Release(ref windows);
                    ComUtilities.Release(ref neighbor);
                    ComUtilities.Release(ref books);
                }
            });
            var discovery = await service.ProcessAsync(new ServiceRequest { Command = "session.discover" });
            RequireSuccess(discovery);
            using var json = JsonDocument.Parse(discovery.Result!);
            Assert.True(json.RootElement.GetProperty("complete").GetBoolean(), discovery.Result);
            var selection = Assert.Single(json.RootElement.GetProperty("workbooks").EnumerateArray(),
                item => item.GetProperty("name").GetString() == Path.GetFileName(neighborPath)).Clone();
            var neighborId = await AttachReviewWorkbookAsync(service, selection);
            sessions[neighborId] = 0;
            var settings = owner.Execute((ctx, _) => (ctx.App.AutomationSecurity, ctx.App.DisplayAlerts));
            var response = await SendSecurityAsync(service, id, "enable-macros",
                new { approved = true, allowReopen = true });
            Assert.False(response.Success);
            Assert.Contains("Another workbook", response.ErrorMessage!, StringComparison.Ordinal);
            Assert.Equal("saved original", await ReadMarkerAsync(service, id));
            Assert.True(owner.Execute((ctx, _) => ctx.Book.Saved));
            Assert.Equal(settings, owner.Execute((ctx, _) => (ctx.App.AutomationSecurity, ctx.App.DisplayAlerts)));
            var neighborNotices = await SendSecurityAsync(service, neighborId, "notices");
            RequireSuccess(neighborNotices);
            using var notices = JsonDocument.Parse(neighborNotices.Result!);
            Assert.Equal(protectedView, notices.RootElement.GetProperty("notices").GetProperty("protectedView").GetBoolean());
            await CloseSessionAsync(service, neighborId, save: false);
            sessions.TryRemove(neighborId, out _);
        });
    }

    [Theory]
    [InlineData("range.get-values")]
    [InlineData("session.save")]
    public async Task SecurityMacros_OpeningCodeClosesWorkbook_ReportsPartialStateAndRestoresSettings(string subsequentCommand)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "self-closing.xlsm");
            var created = await CreateSessionAsync(service, path);
            sessions[created] = 0;
            await WriteMarkerAsync(service, created, "saved original");
            service.SessionManager.GetSession(created)!.Execute((ctx, _) => InstallSecurityProbeMacros(ctx, closeOnOpen: true));
            await CloseSessionAsync(service, created, save: true);
            sessions.TryRemove(created, out _);
            var id = await OpenSessionAsync(service, path);
            sessions[id] = 0;
            var owner = service.SessionManager.GetSession(id)!;
            var originalContext = owner.Execute((ctx, _) => ctx);
            var settings = owner.Execute((ctx, _) => (ctx.App.AutomationSecurity, ctx.App.DisplayAlerts));
            var response = await SendSecurityAsync(service, id, "enable-macros",
                new { approved = true, allowReopen = true });
            Assert.False(response.Success);
            Assert.Contains("original saved workbook closed", response.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(settings, owner.Execute((_, _) =>
                (originalContext.App.AutomationSecurity, originalContext.App.DisplayAlerts)));
            Assert.Equal(0, owner.Execute((_, _) =>
            {
                Excel.Workbooks? books = null;
                try
                {
                    books = originalContext.App.Workbooks;
                    return books.Count;
                }
                finally { ComUtilities.Release(ref books); }
            }));
            var notices = await SendSecurityAsync(service, id, "notices");
            Assert.False(notices.Success);
            Assert.Contains("original saved workbook closed", notices.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
            var subsequent = await service.ProcessAsync(new ServiceRequest
            {
                Command = subsequentCommand,
                SessionId = id,
                Args = subsequentCommand == "range.get-values" ? """{"sheetName":"Sheet1","rangeAddress":"A1"}""" : null
            });
            Assert.False(subsequent.Success);
            Assert.Contains("original saved workbook closed", subsequent.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(nameof(InvalidOperationException), subsequent.ExceptionType);
            await CloseSessionAsync(service, id, save: false);
            sessions.TryRemove(id, out _);
            var recovered = await OpenSessionAsync(service, path);
            sessions[recovered] = 0;
            Assert.Equal("saved original", await ReadMarkerAsync(service, recovered));
        });
    }

    private static Task<ServiceResponse> SendSecurityAsync(ExcelMcpService service, string id, string action, object? args = null) =>
        service.ProcessAsync(new ServiceRequest
        {
            Command = "session." + action,
            SessionId = id,
            Args = args is null ? null : JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions)
        });

    private static void InstallSecurityProbeMacros(ExcelContext ctx, bool closeOnOpen = false)
    {
        Vbe.VBProject? project = null;
        Vbe.VBComponents? components = null;
        Vbe.VBComponent? component = null;
        Vbe.CodeModule? code = null;
        try
        {
            project = ctx.Book.VBProject;
            components = project.VBComponents;
            component = components.Item(ctx.Book.CodeName);
            code = component.CodeModule;
            code.AddFromString(
                "Private Sub Workbook_Open()\r\n" +
                (closeOnOpen ? "ThisWorkbook.Close SaveChanges:=False\r\n"
                    : "ThisWorkbook.Worksheets(1).Range(\"A1\").Value = \"opening-ran\"\r\n") +
                "End Sub\r\n" +
                "Public Sub SecurityMarker()\r\n" +
                "ThisWorkbook.Worksheets(1).Range(\"A1\").Value = \"explicit-ran-\" & CStr(Application.AutomationSecurity)\r\nEnd Sub");
        }
        finally
        {
            ComUtilities.Release(ref code);
            ComUtilities.Release(ref component);
            ComUtilities.Release(ref components);
            ComUtilities.Release(ref project);
        }
    }
}
