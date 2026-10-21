using System.Runtime.ExceptionServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task CloseOwnedSession_WhenExcelCancelsClose_RetainsSessionAndUnsavedWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(
                service, Path.Join(directory, "cancel-owned-close.xlsx"));
            sessions[sessionId] = 0;
            var batch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(sessionId));
            await WriteMarkerAsync(service, sessionId, "keep this edit");

            var closeEventCount = 0;
            Excel.AppEvents_WorkbookBeforeCloseEventHandler? cancelOnce = null;
            batch.Execute((context, _) =>
            {
                cancelOnce = (Excel.Workbook _, ref bool cancel) =>
                {
                    closeEventCount++;
                    cancel = true;
                    context.App.WorkbookBeforeClose -= cancelOnce;
                };
                context.App.WorkbookBeforeClose += cancelOnce;
            });

            var cancelledClose = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = sessionId,
                Args = """{"save":false}"""
            });

            Assert.False(cancelledClose.Success);
            Assert.Contains("did not close", cancelledClose.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, closeEventCount);
            Assert.Same(batch, service.SessionManager.GetSession(sessionId));
            batch.Execute((context, _) =>
            {
                Excel.Workbooks? workbooks = null;
                Excel.Workbook? workbook = null;
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.Range? cell = null;
                try
                {
                    workbooks = context.App.Workbooks;
                    Assert.Equal(1, workbooks.Count);
                    workbook = workbooks[1];
                    Assert.Same(context.Book, workbook);
                    Assert.False(workbook.Saved);
                    sheets = workbook.Worksheets;
                    sheet = (Excel.Worksheet)sheets[1];
                    cell = sheet.Range["A1"];
                    Assert.Equal("keep this edit", cell.Value2);
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                    ComUtilities.Release(ref workbook);
                    ComUtilities.Release(ref workbooks);
                }
            });

            await CloseSessionAsync(service, sessionId, save: false);
            sessions.TryRemove(sessionId, out _);
            Assert.Null(service.SessionManager.GetSession(sessionId));
        });
    }

    [Fact]
    public async Task Discover_WhenOwnedWorkbookIsClosedAndReopenedAtSamePath_AttachesToNewInstance()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "reopened-owned.xlsx");
            var oldSessionId = await CreateSessionAsync(service, path);
            sessions[oldSessionId] = 0;
            var oldBatch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(oldSessionId));
            string neighborName = string.Empty;

            var originalDiscovery = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.discover"
            });
            RequireSuccess(originalDiscovery);
            using var originalJson = JsonDocument.Parse(originalDiscovery.Result!);
            var originalSelection = Assert.Single(
                originalJson.RootElement.GetProperty("workbooks").EnumerateArray(),
                item => item.TryGetProperty("location", out var location)
                    && string.Equals(location.GetString(), path,
                    StringComparison.OrdinalIgnoreCase));
            var originalWorkbookId = originalSelection.GetProperty("workbookId").GetString();
            Assert.Equal(oldSessionId, originalSelection.GetProperty("sessionId").GetString());

            oldBatch.Execute((context, _) =>
            {
                Excel.Workbooks? workbooks = null;
                Excel.Workbook? target = null;
                Excel.Workbook? neighbor = null;
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.Range? cell = null;
                try
                {
                    workbooks = context.App.Workbooks;
                    target = workbooks[1];
                    Assert.Same(context.Book, target);
                    target.Save();
                    neighbor = workbooks.Add();
                    neighborName = neighbor.Name;
                    sheets = neighbor.Worksheets;
                    sheet = (Excel.Worksheet)sheets[1];
                    cell = sheet.Range["A1"];
                    cell.Value2 = "neighbor remains unsaved";
                    Assert.False(neighbor.Saved);

                    target.Close(SaveChanges: false);
                    ComUtilities.Release(ref target);
                    dynamic workbookDispatch = workbooks;
                    target = (Excel.Workbook)workbookDispatch.Open(
                        path, UpdateLinks: 0, ReadOnly: false,
                        IgnoreReadOnlyRecommended: true, Notify: false, AddToMru: false);
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                    ComUtilities.Release(ref neighbor);
                    ComUtilities.Release(ref target);
                    ComUtilities.Release(ref workbooks);
                }
            });

            var discovery = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.discover"
            });
            RequireSuccess(discovery);
            using var discoveryJson = JsonDocument.Parse(discovery.Result!);
            var root = discoveryJson.RootElement;
            var selected = Assert.Single(root.GetProperty("workbooks").EnumerateArray(),
                item => item.TryGetProperty("location", out var location)
                    && string.Equals(location.GetString(), path,
                    StringComparison.OrdinalIgnoreCase));
            Assert.False(selected.TryGetProperty("sessionId", out _));
            var newWorkbookId = selected.GetProperty("workbookId").GetString();
            Assert.False(string.IsNullOrWhiteSpace(newWorkbookId));
            Assert.NotEqual(originalWorkbookId, newWorkbookId);
            Assert.DoesNotContain(root.GetProperty("errors").EnumerateArray(),
                item => item.GetProperty("processId").GetInt32() == oldBatch.ExcelProcessId);
            Assert.DoesNotContain(service.SessionManager.GetActiveSessions(),
                session => session.SessionId == oldSessionId);
            sessions.TryRemove(oldSessionId, out _);

            var attached = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.attach",
                Args = JsonSerializer.Serialize(new { workbookId = newWorkbookId }, ServiceProtocol.JsonOptions)
            });
            RequireSuccess(attached);
            var attachedSessionId = GetSessionId(attached);
            sessions[attachedSessionId] = 0;

            var attachedBatch = Assert.IsAssignableFrom<IExcelBatch>(
                service.SessionManager.GetSession(attachedSessionId));
            attachedBatch.Execute((context, _) =>
            {
                Excel.Workbooks? workbooks = null;
                Excel.Workbook? neighbor = null;
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.Range? cell = null;
                try
                {
                    workbooks = context.App.Workbooks;
                    Assert.Equal(2, workbooks.Count);
                    neighbor = workbooks[neighborName];
                    Assert.False(neighbor.Saved);
                    sheets = neighbor.Worksheets;
                    sheet = (Excel.Worksheet)sheets[1];
                    cell = sheet.Range["A1"];
                    Assert.Equal("neighbor remains unsaved", cell.Value2);
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                    ComUtilities.Release(ref neighbor);
                    ComUtilities.Release(ref workbooks);
                }
            });

            await CloseSessionAsync(service, attachedSessionId, save: false);
            sessions.TryRemove(attachedSessionId, out _);
            var neighborDiscovery = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.discover"
            });
            RequireSuccess(neighborDiscovery);
            using var neighborJson = JsonDocument.Parse(neighborDiscovery.Result!);
            var neighborSelection = Assert.Single(
                neighborJson.RootElement.GetProperty("workbooks").EnumerateArray(),
                item => item.GetProperty("name").GetString() == neighborName);
            Assert.False(neighborSelection.GetProperty("saved").GetBoolean());

            var neighborAttach = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.attach",
                Args = JsonSerializer.Serialize(
                    new { workbookId = neighborSelection.GetProperty("workbookId").GetString() },
                    ServiceProtocol.JsonOptions)
            });
            RequireSuccess(neighborAttach);
            var neighborSessionId = GetSessionId(neighborAttach);
            sessions[neighborSessionId] = 0;
            await CloseSessionAsync(service, neighborSessionId, save: false);
            sessions.TryRemove(neighborSessionId, out _);
        });
    }

    [Fact]
    public async Task DiscoverSelections_TrackRenameAndRejectWorkbookClosedAndReopened()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var owner = new ExcelMcpService();
            var secondOwner = new ExcelMcpService();
            string? ownerSession = null;
            string? secondOwnerSession = null;
            string? attachedSession = null;
            string? managedAttachedSession = null;
            string? managedSession = null;
            try
            {
                managedSession = await CreateSessionAsync(
                    service, Path.Join(directory, "selection-managed.xlsx"));
                sessions[managedSession] = 0;
                var managedBatch = Assert.IsType<ExcelBatch>(
                    service.SessionManager.GetSession(managedSession));
                ownerSession = await CreateSessionAsync(
                    owner, Path.Join(directory, "selection-neighbor.xlsx"));
                secondOwnerSession = await CreateSessionAsync(
                    secondOwner, Path.Join(directory, "selection-second-app.xlsx"));
                var ownerBatch = Assert.IsType<ExcelBatch>(
                    owner.SessionManager.GetSession(ownerSession));
                var secondOwnerBatch = Assert.IsType<ExcelBatch>(
                    secondOwner.SessionManager.GetSession(secondOwnerSession));
                Assert.NotEqual(ownerBatch.ExcelProcessId, secondOwnerBatch.ExcelProcessId);
                var originalPath = Path.Join(directory, "selection-original.xlsx");
                var renamedPath = Path.Join(directory, "selection-renamed.xlsx");
                var managedNeighborPath = Path.Join(directory, "selection-managed-neighbor.xlsx");

                managedBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? workbook = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        workbook = workbooks.Add();
                        workbook.SaveAs(managedNeighborPath, ComInteropConstants.XlOpenXmlWorkbook);
                    }
                    finally
                    {
                        ComUtilities.Release(ref workbook);
                        ComUtilities.Release(ref workbooks);
                    }
                });

                ownerBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? workbook = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        workbook = workbooks.Add();
                        workbook.SaveAs(originalPath, ComInteropConstants.XlOpenXmlWorkbook);
                    }
                    finally
                    {
                        ComUtilities.Release(ref workbook);
                        ComUtilities.Release(ref workbooks);
                    }
                });

                var discovery = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.discover"
                });
                RequireSuccess(discovery);
                using var firstDiscovery = JsonDocument.Parse(discovery.Result!);
                Assert.DoesNotContain(
                    firstDiscovery.RootElement.GetProperty("errors").EnumerateArray(),
                    item => item.GetProperty("processId").GetInt32() == managedBatch.ExcelProcessId);
                var managedSelection = Assert.Single(
                    firstDiscovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                    item => item.TryGetProperty("location", out var location)
                        && location.GetString() == managedBatch.WorkbookPath);
                Assert.Equal(managedSession, managedSelection.GetProperty("sessionId").GetString());
                var managedNeighborSelection = Assert.Single(
                    firstDiscovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                    item => item.TryGetProperty("location", out var location)
                        && location.GetString() == managedNeighborPath);
                Assert.False(managedNeighborSelection.TryGetProperty("sessionId", out _));
                var managedNeighborAttach = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(
                        new { workbookId = managedNeighborSelection.GetProperty("workbookId").GetString() },
                        ServiceProtocol.JsonOptions)
                });
                RequireSuccess(managedNeighborAttach);
                managedAttachedSession = GetSessionId(managedNeighborAttach);
                sessions[managedAttachedSession] = 0;
                var managedNeighborSheets = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "sheet.list",
                    SessionId = managedAttachedSession
                });
                RequireSuccess(managedNeighborSheets);
                await CloseSessionAsync(service, managedAttachedSession, save: false);
                sessions.TryRemove(managedAttachedSession, out _);
                managedAttachedSession = null;
                managedBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? managedWorkbook = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        Assert.Equal(1, workbooks.Count);
                        managedWorkbook = workbooks[1];
                        Assert.Equal("selection-managed.xlsx", managedWorkbook.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref managedWorkbook);
                        ComUtilities.Release(ref workbooks);
                    }
                });
                var firstSelection = Assert.Single(
                    firstDiscovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                    item => item.TryGetProperty("location", out var location)
                        && location.GetString() == originalPath);
                var secondAppSelection = Assert.Single(
                    firstDiscovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                    item => item.GetProperty("processId").GetInt32()
                        == secondOwnerBatch.ExcelProcessId);
                Assert.Equal("selection-second-app.xlsx", secondAppSelection.GetProperty("name").GetString());
                var originalId = firstSelection.GetProperty("workbookId").GetString();
                Assert.False(string.IsNullOrWhiteSpace(originalId));
                var originalAttach = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(
                        new { workbookId = originalId }, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(originalAttach);
                attachedSession = GetSessionId(originalAttach);
                sessions[attachedSession] = 0;

                ownerBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? workbook = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        workbook = workbooks["selection-original.xlsx"];
                        workbook.SaveAs(renamedPath, ComInteropConstants.XlOpenXmlWorkbook);
                    }
                    finally
                    {
                        ComUtilities.Release(ref workbook);
                        ComUtilities.Release(ref workbooks);
                    }
                });

                var renamedSheetList = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "sheet.list",
                    SessionId = attachedSession,
                    Args = JsonSerializer.Serialize(
                        new { filePath = renamedPath }, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(renamedSheetList);

                discovery = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.discover"
                });
                RequireSuccess(discovery);
                using (var renamedDiscovery = JsonDocument.Parse(discovery.Result!))
                {
                    var renamedSelection = Assert.Single(
                        renamedDiscovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                        item => item.TryGetProperty("location", out var location)
                            && location.GetString() == renamedPath);
                    Assert.Equal(originalId, renamedSelection.GetProperty("workbookId").GetString());
                    Assert.Equal(attachedSession, renamedSelection.GetProperty("sessionId").GetString());
                }

                Assert.True(service.SessionManager.TryGetFilePath(attachedSession, out var trackedPath));
                Assert.Equal(renamedPath, trackedPath, ignoreCase: true);
                var sheetList = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "sheet.list",
                    SessionId = attachedSession,
                    Args = JsonSerializer.Serialize(
                        new { filePath = renamedPath }, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(sheetList);

                ownerBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? workbook = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        workbook = workbooks["selection-renamed.xlsx"];
                        workbook.Close(SaveChanges: false);
                        ComUtilities.Release(ref workbook);
                        dynamic workbookDispatch = workbooks;
                        workbook = (Excel.Workbook)workbookDispatch.Open(
                            renamedPath, UpdateLinks: 0, ReadOnly: false,
                            IgnoreReadOnlyRecommended: true, Notify: false, AddToMru: false);
                    }
                    finally
                    {
                        ComUtilities.Release(ref workbook);
                        ComUtilities.Release(ref workbooks);
                    }
                });

                discovery = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.discover"
                });
                RequireSuccess(discovery);
                string? reopenedId;
                using (var reopenedDiscovery = JsonDocument.Parse(discovery.Result!))
                {
                    var reopenedSelection = Assert.Single(
                        reopenedDiscovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                        item => item.TryGetProperty("location", out var location)
                            && location.GetString() == renamedPath);
                    Assert.NotEqual(originalId, reopenedSelection.GetProperty("workbookId").GetString());
                    Assert.False(reopenedSelection.TryGetProperty("sessionId", out _));
                    reopenedId = reopenedSelection.GetProperty("workbookId").GetString();
                }
                Assert.False(string.IsNullOrWhiteSpace(reopenedId));

                Assert.DoesNotContain(
                    service.SessionManager.GetActiveSessions(),
                    session => session.SessionId == attachedSession);
                sessions.TryRemove(attachedSession, out _);
                attachedSession = null;

                var staleAttach = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(
                        new { workbookId = originalId }, ServiceProtocol.JsonOptions)
                });
                Assert.False(staleAttach.Success);
                Assert.True(
                    staleAttach.ErrorMessage!.Contains("unknown open-workbook selection", StringComparison.OrdinalIgnoreCase)
                    || staleAttach.ErrorMessage.Contains("no longer confirmed open", StringComparison.OrdinalIgnoreCase),
                    staleAttach.ErrorMessage);

                ownerBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? reopened = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        Assert.Equal(2, workbooks.Count);
                        reopened = workbooks["selection-renamed.xlsx"];
                        Assert.Equal(renamedPath, reopened.FullName);
                    }
                    finally
                    {
                        ComUtilities.Release(ref reopened);
                        ComUtilities.Release(ref workbooks);
                    }
                });

                var attached = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(
                        new { workbookId = reopenedId }, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(attached);
                attachedSession = GetSessionId(attached);
                sessions[attachedSession] = 0;

                await WriteMarkerAsync(service, attachedSession!, "reopened workbook");
                Assert.Equal("reopened workbook", await ReadMarkerAsync(service, attachedSession!));
                await CloseSessionAsync(service, attachedSession!, save: false);
                sessions.TryRemove(attachedSession!, out _);
                attachedSession = null;

                ownerBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? neighbor = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        Assert.Equal(1, workbooks.Count);
                        neighbor = workbooks[1];
                        Assert.Equal("selection-neighbor.xlsx", neighbor.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref neighbor);
                        ComUtilities.Release(ref workbooks);
                    }
                });
                await CloseSessionAsync(service, managedSession, save: false);
                sessions.TryRemove(managedSession, out _);
                managedSession = null;
                service.SessionManager.Dispose();
                await CloseSessionAsync(owner, ownerSession, save: false);
                ownerSession = null;
                await CloseSessionAsync(secondOwner, secondOwnerSession, save: false);
                secondOwnerSession = null;
            }
            finally
            {
                try
                {
                    service.SessionManager.Dispose();
                }
                finally
                {
                    try
                    {
                        if (ownerSession is not null
                            && owner.SessionManager.GetActiveSessions().Any(
                                session => session.SessionId == ownerSession))
                        {
                            await CloseSessionAsync(owner, ownerSession, save: false);
                        }
                    }
                    finally
                    {
                        try
                        {
                            if (secondOwnerSession is not null
                                && secondOwner.SessionManager.GetActiveSessions().Any(
                                    session => session.SessionId == secondOwnerSession))
                            {
                                await CloseSessionAsync(secondOwner, secondOwnerSession, save: false);
                            }
                        }
                        finally
                        {
                            owner.Dispose();
                            secondOwner.Dispose();
                        }
                    }
                }
            }
        });
    }

    [Fact]
    public async Task DiscoverAndAttach_UnsavedWorkbook_ClosesOnlySelectedWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var owner = new ExcelMcpService();
            string? ownerSession = null;
            string? attachedSession = null;
            string? shutdownWorkbookName = null;
            Exception? failure = null;
            try
            {
                ownerSession = await CreateSessionAsync(
                    owner, Path.Join(directory, "neighbor.xlsx"));
                await WriteMarkerAsync(owner, ownerSession, "neighbor unsaved");

                var ownerBatch = Assert.IsType<ExcelBatch>(
                    owner.SessionManager.GetSession(ownerSession));
                var selectedWorkbookName = ownerBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? workbook = null;
                    Excel.Sheets? sheets = null;
                    Excel.Worksheet? sheet = null;
                    Excel.Range? cell = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        workbook = workbooks.Add();
                        sheets = workbook.Worksheets;
                        sheet = (Excel.Worksheet)sheets[1];
                        cell = sheet.Range["A1"];
                        cell.Value2 = "selected unsaved";
                        Assert.False(workbook.Saved);
                        return workbook.Name;
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                        ComUtilities.Release(ref sheet);
                        ComUtilities.Release(ref sheets);
                        ComUtilities.Release(ref workbook);
                        ComUtilities.Release(ref workbooks);
                    }
                });

                var discovery = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.discover"
                });
                RequireSuccess(discovery);
                using var discoveryJson = JsonDocument.Parse(discovery.Result!);
                var root = discoveryJson.RootElement;
                var target = Assert.Single(root.GetProperty("workbooks").EnumerateArray(),
                    item => item.GetProperty("processId").GetInt32() == ownerBatch.ExcelProcessId
                        && item.GetProperty("name").GetString() == selectedWorkbookName);
                Assert.False(target.GetProperty("saved").GetBoolean());
                Assert.False(target.TryGetProperty("location", out _));
                var discoveredAutoSave = target.TryGetProperty("autoSaveOn", out var autoSave)
                    ? autoSave.GetBoolean()
                    : (bool?)null;
                var workbookId = target.GetProperty("workbookId").GetString();
                Assert.False(string.IsNullOrWhiteSpace(workbookId));

                var attached = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(new { workbookId }, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(attached);
                attachedSession = GetSessionId(attached);
                sessions[attachedSession] = 0;
                using (var attachedJson = JsonDocument.Parse(attached.Result!))
                {
                    var attachedResult = attachedJson.RootElement;
                    var attachedAutoSave = attachedResult.TryGetProperty("autoSaveOn", out var autoSaveOn)
                        ? autoSaveOn.GetBoolean()
                        : (bool?)null;
                    Assert.Equal(discoveredAutoSave, attachedAutoSave);
                    Assert.False(attachedResult.TryGetProperty("filePath", out _));
                    var warning = attachedResult.GetProperty("warning").GetString();
                    Assert.NotNull(warning);
                    Assert.Contains(discoveredAutoSave switch
                    {
                        true => "AutoSave is enabled",
                        false => "AutoSave is off",
                        null => "AutoSave status is unavailable"
                    }, warning, StringComparison.Ordinal);
                }

                discovery = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.discover"
                });
                RequireSuccess(discovery);
                using (var attachedRediscovery = JsonDocument.Parse(discovery.Result!))
                {
                    var rediscovered = Assert.Single(
                        attachedRediscovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                        item => item.GetProperty("processId").GetInt32() == ownerBatch.ExcelProcessId
                            && item.GetProperty("name").GetString() == selectedWorkbookName);
                    Assert.False(rediscovered.GetProperty("saved").GetBoolean());
                    Assert.Equal(attachedSession, rediscovered.GetProperty("sessionId").GetString());
                    Assert.DoesNotContain(
                        attachedRediscovery.RootElement.GetProperty("errors").EnumerateArray(),
                        item => item.GetProperty("processId").GetInt32() == ownerBatch.ExcelProcessId);
                }

                var liveAutoSave = ownerBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? selected = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        selected = workbooks[selectedWorkbookName];
                        return ExcelCapabilities.ReadAutoSaveStatus(() => selected.AutoSaveOn);
                    }
                    finally
                    {
                        ComUtilities.Release(ref selected);
                        ComUtilities.Release(ref workbooks);
                    }
                });
                Assert.Equal(discoveredAutoSave, liveAutoSave);

                var attachedBatch = Assert.IsAssignableFrom<IExcelBatch>(
                    service.SessionManager.GetSession(attachedSession));
                var closeEventCount = 0;
                Excel.AppEvents_WorkbookBeforeCloseEventHandler cancelClose =
                    (Excel.Workbook _, ref bool cancel) =>
                    {
                        closeEventCount++;
                        cancel = true;
                    };
                attachedBatch.Execute((context, _) => context.App.WorkbookBeforeClose += cancelClose);
                try
                {
                    var cancelledClose = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.close",
                        SessionId = attachedSession,
                        Args = """{"save":false}"""
                    });
                    Assert.False(cancelledClose.Success);
                    Assert.Contains("did not close", cancelledClose.ErrorMessage, StringComparison.OrdinalIgnoreCase);
                    Assert.True(closeEventCount > 0, "Excel did not invoke the before-close event.");
                    Assert.Same(attachedBatch, service.SessionManager.GetSession(attachedSession));
                    ownerBatch.Execute((context, _) =>
                    {
                        Excel.Workbooks? workbooks = null;
                        Excel.Workbook? selected = null;
                        Excel.Sheets? sheets = null;
                        Excel.Worksheet? sheet = null;
                        Excel.Range? cell = null;
                        try
                        {
                            workbooks = context.App.Workbooks;
                            Assert.Equal(2, workbooks.Count);
                            selected = workbooks[selectedWorkbookName];
                            Assert.False(selected.Saved);
                            sheets = selected.Worksheets;
                            sheet = (Excel.Worksheet)sheets[1];
                            cell = sheet.Range["A1"];
                            Assert.Equal("selected unsaved", cell.Value2);
                        }
                        finally
                        {
                            ComUtilities.Release(ref cell);
                            ComUtilities.Release(ref sheet);
                            ComUtilities.Release(ref sheets);
                            ComUtilities.Release(ref selected);
                            ComUtilities.Release(ref workbooks);
                        }
                    });
                }
                finally
                {
                    attachedBatch.Execute((context, _) => context.App.WorkbookBeforeClose -= cancelClose);
                }

                var sessionList = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.list"
                });
                RequireSuccess(sessionList);
                using (var sessionListJson = JsonDocument.Parse(sessionList.Result!))
                {
                    var attachedEntry = Assert.Single(
                        sessionListJson.RootElement.GetProperty("sessions").EnumerateArray(),
                        item => item.GetProperty("sessionId").GetString() == attachedSession);
                    Assert.True(attachedEntry.GetProperty("isAttached").GetBoolean());
                    Assert.Equal(selectedWorkbookName, attachedEntry.GetProperty("workbookName").GetString());
                    Assert.False(attachedEntry.TryGetProperty("filePath", out _));
                }

                using (var cancellationStarted = new ManualResetEventSlim())
                using (var cancellationRelease = new ManualResetEventSlim())
                using (var cancellationFinished = new ManualResetEventSlim())
                using (var cancellation = new CancellationTokenSource())
                {
                    var pendingOperation = Task.Run(() => attachedBatch.Execute((_, _) =>
                    {
                        cancellationStarted.Set();
                        try
                        {
                            if (!cancellationRelease.Wait(TimeSpan.FromSeconds(5)))
                                throw new TimeoutException("Test did not release the attached operation.");
                        }
                        finally
                        {
                            cancellationFinished.Set();
                        }
                    }, cancellation.Token));
                    Assert.True(cancellationStarted.Wait(TimeSpan.FromSeconds(5)));
                    cancellation.CancelAfter(TimeSpan.FromMilliseconds(100));
                    await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await pendingOperation);
                    cancellationRelease.Set();
                    Assert.True(cancellationFinished.Wait(TimeSpan.FromSeconds(5)));
                    var refreshState = WorkbookRefreshState.Unknown;
                    var readinessDeadline = DateTime.UtcNow + TimeSpan.FromSeconds(5);
                    while (refreshState != WorkbookRefreshState.Ready
                        && DateTime.UtcNow < readinessDeadline)
                    {
                        refreshState =
                            Assert.IsAssignableFrom<IExcelBatchRefreshState>(attachedBatch).GetRefreshState();
                        if (refreshState != WorkbookRefreshState.Ready)
                            await Task.Delay(TimeSpan.FromMilliseconds(50));
                    }
                    Assert.Equal(WorkbookRefreshState.Ready, refreshState);
                }

                await WriteMarkerAsync(service, attachedSession, "attached edit");
                Assert.Equal("attached edit", await ReadMarkerAsync(service, attachedSession));
                Assert.Equal("neighbor unsaved", await ReadMarkerAsync(owner, ownerSession));

                await CloseSessionAsync(service, attachedSession, save: false);
                sessions.TryRemove(attachedSession, out _);
                Assert.DoesNotContain(
                    service.SessionManager.GetActiveSessions(),
                    session => session.SessionId == attachedSession);

                var remaining = ownerBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? neighbor = null;
                    Excel.Sheets? sheets = null;
                    Excel.Worksheet? sheet = null;
                    Excel.Range? cell = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        Assert.Equal(1, workbooks.Count);
                        neighbor = workbooks[1];
                        Assert.Equal("neighbor.xlsx", neighbor.Name);
                        sheets = neighbor.Worksheets;
                        sheet = (Excel.Worksheet)sheets[1];
                        cell = sheet.Range["A1"];
                        Assert.Equal("neighbor unsaved", cell.Value2);
                        Assert.False(neighbor.Saved);
                        return workbooks.Count;
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                        ComUtilities.Release(ref sheet);
                        ComUtilities.Release(ref sheets);
                        ComUtilities.Release(ref neighbor);
                        ComUtilities.Release(ref workbooks);
                    }
                });
                Assert.Equal(1, remaining);

                shutdownWorkbookName = ownerBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? workbook = null;
                    Excel.Sheets? sheets = null;
                    Excel.Worksheet? sheet = null;
                    Excel.Range? cell = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        workbook = workbooks.Add();
                        sheets = workbook.Worksheets;
                        sheet = (Excel.Worksheet)sheets[1];
                        cell = sheet.Range["A1"];
                        cell.Value2 = "shutdown target";
                        Assert.False(workbook.Saved);
                        return workbook.Name;
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                        ComUtilities.Release(ref sheet);
                        ComUtilities.Release(ref sheets);
                        ComUtilities.Release(ref workbook);
                        ComUtilities.Release(ref workbooks);
                    }
                });

                discovery = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.discover"
                });
                RequireSuccess(discovery);
                using var shutdownDiscoveryJson = JsonDocument.Parse(discovery.Result!);
                var shutdownTarget = Assert.Single(
                    shutdownDiscoveryJson.RootElement.GetProperty("workbooks").EnumerateArray(),
                    item => item.GetProperty("processId").GetInt32() == ownerBatch.ExcelProcessId
                        && item.GetProperty("name").GetString() == shutdownWorkbookName);
                var shutdownWorkbookId = shutdownTarget.GetProperty("workbookId").GetString();
                var shutdownAttach = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(
                        new { workbookId = shutdownWorkbookId }, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(shutdownAttach);
                var shutdownSession = GetSessionId(shutdownAttach);
                sessions[shutdownSession] = 0;
                await WriteMarkerAsync(service, shutdownSession, "shutdown unsaved");

                sessions.TryRemove(shutdownSession, out _);
                service.SessionManager.Dispose();

                var openAfterDetach = ownerBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? neighbor = null;
                    Excel.Workbook? shutdownTargetBook = null;
                    Excel.Sheets? sheets = null;
                    Excel.Worksheet? sheet = null;
                    Excel.Range? cell = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        Assert.Equal(2, workbooks.Count);
                        neighbor = workbooks["neighbor.xlsx"];
                        shutdownTargetBook = workbooks[shutdownWorkbookName];
                        Assert.False(neighbor.Saved);
                        Assert.False(shutdownTargetBook.Saved);
                        sheets = shutdownTargetBook.Worksheets;
                        sheet = (Excel.Worksheet)sheets[1];
                        cell = sheet.Range["A1"];
                        Assert.Equal("shutdown unsaved", cell.Value2);
                        return workbooks.Count;
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                        ComUtilities.Release(ref sheet);
                        ComUtilities.Release(ref sheets);
                        ComUtilities.Release(ref shutdownTargetBook);
                        ComUtilities.Release(ref neighbor);
                        ComUtilities.Release(ref workbooks);
                    }
                });
                Assert.Equal(2, openAfterDetach);

                var cleanupService = new ExcelMcpService();
                Exception? cleanupFailure = null;
                try
                {
                    discovery = await cleanupService.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.discover"
                    });
                    RequireSuccess(discovery);
                    using var cleanupDiscoveryJson = JsonDocument.Parse(discovery.Result!);
                    var cleanupTarget = Assert.Single(
                        cleanupDiscoveryJson.RootElement.GetProperty("workbooks").EnumerateArray(),
                        item => item.GetProperty("processId").GetInt32() == ownerBatch.ExcelProcessId
                            && item.GetProperty("name").GetString() == shutdownWorkbookName);
                    var cleanupAttach = await cleanupService.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.attach",
                        Args = JsonSerializer.Serialize(
                            new { workbookId = cleanupTarget.GetProperty("workbookId").GetString() },
                            ServiceProtocol.JsonOptions)
                    });
                    RequireSuccess(cleanupAttach);
                    var cleanupSession = GetSessionId(cleanupAttach);
                    await CloseSessionAsync(cleanupService, cleanupSession, save: false);
                }
                catch (Exception ex)
                {
                    cleanupFailure = ex;
                }
                finally
                {
                    try
                    {
                        cleanupService.Dispose();
                    }
                    catch (Exception ex)
                    {
                        cleanupFailure = PersistentServiceCleanupFailures.Combine(cleanupFailure, ex);
                    }
                }
                if (cleanupFailure is not null)
                    ExceptionDispatchInfo.Capture(cleanupFailure).Throw();

                await CloseSessionAsync(owner, ownerSession, save: false);
                ownerSession = null;
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            finally
            {
                try
                {
                    service.SessionManager.Dispose();
                }
                catch (Exception ex)
                {
                    failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                }
                try
                {
                    if (ownerSession is not null
                        && owner.SessionManager.GetActiveSessions().Any(
                            session => session.SessionId == ownerSession))
                    {
                        await CloseSessionAsync(owner, ownerSession, save: false);
                    }
                }
                catch (Exception ex)
                {
                    failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                }
                try
                {
                    owner.Dispose();
                }
                catch (Exception ex)
                {
                    failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                }
            }
            if (failure is not null)
                ExceptionDispatchInfo.Capture(failure).Throw();
        });
    }
}
