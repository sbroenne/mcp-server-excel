using System.Collections.Concurrent;
using System.Runtime.ExceptionServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    public async Task CloseOrShutdown_AfterWorkbookClosedInExcel_PreservesOnlyNeighbor(
        int exitMode, bool keepNeighbor)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            ExcelBatch? owned = null;
            await RunWithCleanupAsync(async (observer, observerDirectory, observedSessions) =>
            {
                var path = Path.Join(directory, "closed-in-excel.xlsx");
                var ownedId = await CreateSessionAsync(service, path);
                sessions[ownedId] = 0;
                owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
                var neighborName = keepNeighbor
                    ? owned.Execute((context, _) => AddReviewWorkbook(context.App, "keep neighboring edits"))
                    : null;
                await DiscoverReviewWorkbookAsync(service, Path.GetFileName(path));
                owned.Execute((context, _) => context.Book.Close(SaveChanges: false));
                Assert.True(owned.IsExcelProcessAlive());

                if (exitMode == 2)
                {
                    service.SessionManager.Dispose();
                }
                else
                {
                    if (exitMode == 1)
                    {
                        var saveAttempt = await service.ProcessAsync(new ServiceRequest
                        {
                            Command = "session.close",
                            SessionId = ownedId,
                            Args = JsonSerializer.Serialize(new { save = true }, ServiceProtocol.JsonOptions)
                        });
                        Assert.False(saveAttempt.Success);
                        Assert.Contains("Nothing was saved", saveAttempt.ErrorMessage, StringComparison.Ordinal);
                        Assert.Equal(ownedId, Assert.Single(service.SessionManager.ActiveSessionIds));
                    }
                    await CloseSessionAsync(service, ownedId, save: false);
                }
                sessions.TryRemove(ownedId, out _);

                Assert.Empty(service.SessionManager.ActiveSessionIds);
                Assert.Equal(keepNeighbor, owned.IsExcelProcessAlive());
                if (neighborName is null)
                    return;
                var ownedIdentity = Assert.IsType<ExcelProcessIdentity>(owned.ProcessIdentity);
                var visibleBeforeAttach = IsExcelWindowVisible(ownedIdentity, owned.Logger);

                var neighborId = await AttachReviewWorkbookAsync(observer,
                    await DiscoverReviewWorkbookAsync(observer, neighborName));
                observedSessions[neighborId] = 0;
                Assert.True(visibleBeforeAttach, "Excel left open with another workbook must be visible.");
                var neighbor = Assert.IsType<AttachedWorkbookBatch>(
                    observer.SessionManager.GetSession(neighborId));
                Assert.Equal(owned.ExcelProcessId, neighbor.ExcelProcessId);
                Assert.Equal("keep neighboring edits", await ReadMarkerAsync(observer, neighborId));
                neighbor.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        Assert.Equal(1, workbooks.Count);
                        Assert.True(context.App.DisplayAlerts);
                        Assert.False(context.Book.Saved);
                        Assert.Equal(string.Empty, context.Book.Path);
                    }
                    finally
                    {
                        ComUtilities.Release(ref workbooks);
                    }
                });
            });
            Assert.NotNull(owned);
            Assert.False(owned.IsExcelProcessAlive());
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Open_AfterStatusHidesExitedExcel_ReclaimsPath(bool validateFirst)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "reopen-after-exit.xlsx");
            var createdId = await CreateSessionAsync(service, path);
            sessions[createdId] = 0;
            await WriteMarkerAsync(service, createdId, "saved before Excel exited");
            await CloseSessionAsync(service, createdId, save: true);
            sessions.TryRemove(createdId, out _);

            var oldId = await OpenSessionAsync(service, path);
            sessions[oldId] = 0;
            Assert.Equal("saved before Excel exited", await ReadMarkerAsync(service, oldId));
            var oldBatch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(oldId));

            // Close through the product without informing the session manager, as with an external Excel exit.
            oldBatch.Close();
            Assert.False(oldBatch.IsExcelProcessAlive());
            sessions.TryRemove(oldId, out _);

            var list = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            RequireSuccess(list);
            using (var result = JsonDocument.Parse(list.Result!))
                Assert.Empty(result.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Equal(0, service.SessionCount);
            Assert.Equal(oldId, Assert.Single(service.SessionManager.ActiveSessionIds));

            if (validateFirst)
            {
                var validation = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.test",
                    Args = JsonSerializer.Serialize(new { filePath = path }, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(validation);
                using var result = JsonDocument.Parse(validation.Result!);
                Assert.True(result.RootElement.GetProperty("canOpen").GetBoolean(), validation.Result);
                Assert.Empty(service.SessionManager.ActiveSessionIds);
            }

            var reopenedId = await OpenSessionAsync(service, path);
            sessions[reopenedId] = 0;
            Assert.NotEqual(oldId, reopenedId);
            Assert.Equal(reopenedId, Assert.Single(service.SessionManager.ActiveSessionIds));
            Assert.Equal("saved before Excel exited", await ReadMarkerAsync(service, reopenedId));

            var reopenedList = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            RequireSuccess(reopenedList);
            using var reopenedResult = JsonDocument.Parse(reopenedList.Result!);
            var reopenedSession = Assert.Single(reopenedResult.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Equal(reopenedId, reopenedSession.GetProperty("sessionId").GetString());
            Assert.Equal(path, reopenedSession.GetProperty("filePath").GetString(), ignoreCase: true);
        });
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task DiscoveryCleanup_IncompleteScan_ClosesKnownWorkbookAndReportsError(bool attachFirst)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var ownedId = await CreateSessionAsync(service, Path.Join(directory, "cleanup-owner.xlsx"));
            sessions[ownedId] = 0;
            var owned = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(ownedId));
            var name = owned.Execute((context, _) => AddReviewWorkbook(context.App, "cleanup selected edits"));
            Exception? failure = null;
            try
            {
                string? attachedId = null;
                if (attachFirst)
                {
                    attachedId = await AttachReviewWorkbookAsync(service,
                        await DiscoverReviewWorkbookAsync(service, name));
                    sessions[attachedId] = 0;
                }
                await WriteMarkerAsync(service, ownedId, "keep neighboring edits");
                var discovery = await service.ProcessAsync(new ServiceRequest { Command = "session.discover" });
                RequireSuccess(discovery);
                var result = System.Text.Json.Nodes.JsonNode.Parse(discovery.Result!)!;
                result["complete"] = false;
                result["errors"] = System.Text.Json.Nodes.JsonNode.Parse(
                    """[{"processId":-1,"message":"Synthetic unrelated inspection failure"}]""");
                var incompleteDiscovery = new ServiceResponse { Success = true, Result = result.ToJsonString() };
                var scanFailure = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                    CloseDiscoveredWorkbooksAsync(service, sessions, () => Task.FromResult(incompleteDiscovery),
                        owned.ExcelProcessId!.Value, (null, name)));
                Assert.Contains("Synthetic unrelated inspection failure", scanFailure.Message, StringComparison.Ordinal);
                if (attachedId is not null)
                {
                    Assert.Null(service.SessionManager.GetSession(attachedId));
                    Assert.False(sessions.ContainsKey(attachedId));
                }
                Assert.Single(sessions);
                owned.Execute((context, _) =>
                {
                    Excel.Workbooks? books = null;
                    try
                    {
                        books = context.App.Workbooks;
                        Assert.Equal(1, books.Count);
                        Assert.False(context.Book.Saved);
                    }
                    finally
                    {
                        ComUtilities.Release(ref books);
                    }
                });
                Assert.Equal("keep neighboring edits", await ReadMarkerAsync(service, ownedId));
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            finally
            {
                try
                {
                    await CloseDiscoveredWorkbooksAsync(
                        service, sessions, owned.ExcelProcessId!.Value, (null, name));
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
                    Assert.Equal("keep this edit", ((object)cell.Value2));
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
            string? neighborName = null;
            string? reopenedName = null;
            int? processId = null;
            Exception? failure = null;

            try
            {
                var oldSessionId = await CreateSessionAsync(service, path);
                sessions[oldSessionId] = 0;
                var oldBatch = Assert.IsType<ExcelBatch>(service.SessionManager.GetSession(oldSessionId));
                processId = oldBatch.ExcelProcessId;
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
                        target = workbooks.Open(
                            path, UpdateLinks: 0, ReadOnly: false,
                            IgnoreReadOnlyRecommended: true, Notify: false, AddToMru: false);
                        reopenedName = target.Name;
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
                        Assert.Equal("neighbor remains unsaved", ((object)cell.Value2));
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
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            finally
            {
                if (processId is int ownerProcessId)
                {
                    try
                    {
                        var workbooks = new List<(string? Location, string? Name)>();
                        if (!string.IsNullOrWhiteSpace(neighborName))
                            workbooks.Add((null, neighborName));
                        workbooks.Add((path, reopenedName));
                        await CloseDiscoveredWorkbooksAsync(
                            service, sessions, ownerProcessId, workbooks.ToArray());
                    }
                    catch (Exception ex)
                    {
                        failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                    }
                }
            }

            if (failure is not null)
                ExceptionDispatchInfo.Capture(failure).Throw();
        });
    }

    [Fact]
    public async Task DiscoverSelections_TrackRenameAndRejectWorkbookClosedAndReopened()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var owner = new ExcelMcpService();
            var secondOwner = new ExcelMcpService();
            var originalPath = Path.Join(directory, "selection-original.xlsx");
            var renamedPath = Path.Join(directory, "selection-renamed.xlsx");
            var managedNeighborPath = Path.Join(directory, "selection-managed-neighbor.xlsx");
            string? ownerSession = null;
            string? secondOwnerSession = null;
            string? attachedSession = null;
            string? managedAttachedSession = null;
            string? managedSession = null;
            string? managedNeighborName = null;
            string? ownerWorkbookName = null;
            int? managedProcessIdForCleanup = null;
            int? ownerProcessIdForCleanup = null;
            var serviceManagerDisposed = false;
            Exception? failure = null;
            try
            {
                managedSession = await CreateSessionAsync(
                    service, Path.Join(directory, "selection-managed.xlsx"));
                sessions[managedSession] = 0;
                var managedBatch = Assert.IsType<ExcelBatch>(
                    service.SessionManager.GetSession(managedSession));
                managedProcessIdForCleanup = managedBatch.ExcelProcessId;
                ownerSession = await CreateSessionAsync(
                    owner, Path.Join(directory, "selection-neighbor.xlsx"));
                secondOwnerSession = await CreateSessionAsync(
                    secondOwner, Path.Join(directory, "selection-second-app.xlsx"));
                var ownerBatch = Assert.IsType<ExcelBatch>(
                    owner.SessionManager.GetSession(ownerSession));
                ownerProcessIdForCleanup = ownerBatch.ExcelProcessId;
                var secondOwnerBatch = Assert.IsType<ExcelBatch>(
                    secondOwner.SessionManager.GetSession(secondOwnerSession));
                Assert.NotEqual(ownerBatch.ExcelProcessId, secondOwnerBatch.ExcelProcessId);

                managedBatch.Execute((context, _) =>
                {
                    Excel.Workbooks? workbooks = null;
                    Excel.Workbook? workbook = null;
                    try
                    {
                        workbooks = context.App.Workbooks;
                        workbook = workbooks.Add();
                        managedNeighborName = workbook.Name;
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
                        ownerWorkbookName = workbook.Name;
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
                        workbook = workbooks.Open(
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
                serviceManagerDisposed = true;
                service.SessionManager.Dispose();
                await CloseSessionAsync(owner, ownerSession, save: false);
                ownerSession = null;
                await CloseSessionAsync(secondOwner, secondOwnerSession, save: false);
                secondOwnerSession = null;
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            finally
            {
                foreach (var sessionId in new[] { managedAttachedSession, attachedSession })
                {
                    if (serviceManagerDisposed || sessionId is null)
                        continue;

                    try
                    {
                        if (service.SessionManager.GetActiveSessions().Any(
                                session => session.SessionId == sessionId))
                        {
                            await CloseSessionAsync(service, sessionId, save: false);
                            sessions.TryRemove(sessionId, out _);
                        }
                    }
                    catch (Exception ex)
                    {
                        failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                    }
                }

                if (!serviceManagerDisposed)
                {
                    if (ownerProcessIdForCleanup is int ownerProcessId)
                    {
                        try
                        {
                            await CloseDiscoveredWorkbooksAsync(
                                service,
                                sessions,
                                ownerProcessId,
                                (originalPath, null),
                                (renamedPath, null),
                                (null, ownerWorkbookName));
                        }
                        catch (Exception ex)
                        {
                            failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                        }
                    }

                    if (managedProcessIdForCleanup is int managedProcessId)
                    {
                        try
                        {
                            await CloseDiscoveredWorkbooksAsync(
                                service,
                                sessions,
                                managedProcessId,
                                (managedNeighborPath, null),
                                (null, managedNeighborName));
                        }
                        catch (Exception ex)
                        {
                            failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                        }
                    }
                }

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
                    if (secondOwnerSession is not null
                        && secondOwner.SessionManager.GetActiveSessions().Any(
                            session => session.SessionId == secondOwnerSession))
                    {
                        await CloseSessionAsync(secondOwner, secondOwnerSession, save: false);
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

                try
                {
                    secondOwner.Dispose();
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

    [Fact]
    public async Task DiscoverAndAttach_UnsavedWorkbook_ClosesOnlySelectedWorkbook()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var owner = new ExcelMcpService();
            string? ownerSession = null;
            string? attachedSession = null;
            string? shutdownSession = null;
            string? shutdownWorkbookNameForCleanup = null;
            string? selectedWorkbookNameForCleanup = null;
            int? ownerProcessIdForCleanup = null;
            var serviceManagerDisposed = false;
            Exception? failure = null;
            try
            {
                ownerSession = await CreateSessionAsync(
                    owner, Path.Join(directory, "neighbor.xlsx"));
                await WriteMarkerAsync(owner, ownerSession, "neighbor unsaved");

                var ownerBatch = Assert.IsType<ExcelBatch>(
                    owner.SessionManager.GetSession(ownerSession));
                ownerProcessIdForCleanup = ownerBatch.ExcelProcessId;
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
                        selectedWorkbookNameForCleanup = workbook.Name;
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
                selectedWorkbookNameForCleanup = selectedWorkbookName;

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
                            Assert.Equal("selected unsaved", ((object)cell.Value2));
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
                        Assert.Equal("neighbor unsaved", ((object)cell.Value2));
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

                var createdShutdownWorkbookName = ownerBatch.Execute((context, _) =>
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
                        shutdownWorkbookNameForCleanup = workbook.Name;
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
                        && item.GetProperty("name").GetString() == createdShutdownWorkbookName);
                var shutdownWorkbookId = shutdownTarget.GetProperty("workbookId").GetString();
                var shutdownAttach = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "session.attach",
                    Args = JsonSerializer.Serialize(
                        new { workbookId = shutdownWorkbookId }, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(shutdownAttach);
                var createdShutdownSession = GetSessionId(shutdownAttach);
                shutdownSession = createdShutdownSession;
                sessions[createdShutdownSession] = 0;
                await WriteMarkerAsync(service, createdShutdownSession, "shutdown unsaved");

                sessions.TryRemove(createdShutdownSession, out _);
                serviceManagerDisposed = true;
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
                        shutdownTargetBook = workbooks[createdShutdownWorkbookName];
                        Assert.False(neighbor.Saved);
                        Assert.False(shutdownTargetBook.Saved);
                        sheets = shutdownTargetBook.Worksheets;
                        sheet = (Excel.Worksheet)sheets[1];
                        cell = sheet.Range["A1"];
                        Assert.Equal("shutdown unsaved", ((object)cell.Value2));
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
                            && item.GetProperty("name").GetString() == createdShutdownWorkbookName);
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
                foreach (var sessionId in new[] { attachedSession, shutdownSession })
                {
                    if (serviceManagerDisposed || sessionId is null)
                        continue;

                    try
                    {
                        if (service.SessionManager.GetActiveSessions().Any(
                                session => session.SessionId == sessionId))
                        {
                            await CloseSessionAsync(service, sessionId, save: false);
                            sessions.TryRemove(sessionId, out _);
                        }
                    }
                    catch (Exception ex)
                    {
                        failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                    }
                }

                if (!serviceManagerDisposed && ownerProcessIdForCleanup is int ownerProcessId)
                {
                    try
                    {
                        await CloseDiscoveredWorkbooksAsync(
                            service,
                            sessions,
                            ownerProcessId,
                            (null, selectedWorkbookNameForCleanup),
                            (null, shutdownWorkbookNameForCleanup));
                    }
                    catch (Exception ex)
                    {
                        failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                    }
                }
                else if (ownerProcessIdForCleanup is int detachedOwnerProcessId)
                {
                    var cleanupService = new ExcelMcpService();
                    var cleanupSessions = new ConcurrentDictionary<string, byte>();
                    try
                    {
                        await CloseDiscoveredWorkbooksAsync(
                            cleanupService,
                            cleanupSessions,
                            detachedOwnerProcessId,
                            (null, selectedWorkbookNameForCleanup),
                            (null, shutdownWorkbookNameForCleanup));
                    }
                    catch (Exception ex)
                    {
                        failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                    }
                    finally
                    {
                        foreach (var sessionId in cleanupSessions.Keys)
                        {
                            try
                            {
                                await CloseSessionAsync(cleanupService, sessionId, save: false);
                            }
                            catch (Exception ex)
                            {
                                failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                            }
                        }

                        try
                        {
                            cleanupService.Dispose();
                        }
                        catch (Exception ex)
                        {
                            failure = PersistentServiceCleanupFailures.Combine(failure, ex);
                        }
                    }
                }

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

    private static Task CloseDiscoveredWorkbooksAsync(
        ExcelMcpService service,
        ConcurrentDictionary<string, byte> sessions,
        int processId,
        params (string? Location, string? Name)[] identities) =>
        CloseDiscoveredWorkbooksAsync(service, sessions,
            () => service.ProcessAsync(new ServiceRequest { Command = "session.discover" }),
            processId, identities);

    private static async Task CloseDiscoveredWorkbooksAsync(
        ExcelMcpService service,
        ConcurrentDictionary<string, byte> sessions,
        Func<Task<ServiceResponse>> discover,
        int processId,
        params (string? Location, string? Name)[] identities)
    {
        Exception? failure = null;
        foreach (var identity in identities)
        {
            if (string.IsNullOrWhiteSpace(identity.Location)
                && string.IsNullOrWhiteSpace(identity.Name))
            {
                continue;
            }

            try
            {
                var discovery = await discover();
                RequireSuccess(discovery);
                using var result = JsonDocument.Parse(discovery.Result!);
                if (!result.RootElement.GetProperty("complete").GetBoolean())
                {
                    failure = PersistentServiceCleanupFailures.Combine(failure, new InvalidOperationException(
                        $"Cannot confirm cleanup of test-owned workbooks in Excel process {processId}: " +
                        $"discovery was incomplete. Errors: {result.RootElement.GetProperty("errors").GetRawText()}"));
                }

                var matches = result.RootElement.GetProperty("workbooks").EnumerateArray()
                    .Where(item =>
                    {
                        if (item.GetProperty("processId").GetInt32() != processId)
                            return false;

                        var locationMatches = !string.IsNullOrWhiteSpace(identity.Location)
                            && item.TryGetProperty("location", out var location)
                            && string.Equals(
                                location.GetString(), identity.Location, StringComparison.OrdinalIgnoreCase);
                        var nameMatches = !string.IsNullOrWhiteSpace(identity.Name)
                            && string.Equals(
                                item.GetProperty("name").GetString(), identity.Name, StringComparison.Ordinal);
                        return locationMatches || nameMatches;
                    })
                    .ToArray();
                if (matches.Length == 0)
                    continue;
                if (matches.Length != 1)
                    throw new InvalidOperationException(
                        $"Expected one test-owned workbook match in Excel process {processId}, found {matches.Length}.");

                var selected = matches[0];
                string? sessionId = null;
                if (selected.TryGetProperty("sessionId", out var sessionProperty))
                {
                    var discoveredSessionId = sessionProperty.GetString();
                    if (!string.IsNullOrWhiteSpace(discoveredSessionId)
                        && service.SessionManager.GetActiveSessions().Any(
                            session => session.SessionId == discoveredSessionId))
                    {
                        sessionId = discoveredSessionId;
                    }
                }

                if (sessionId is null)
                {
                    var attach = await service.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.attach",
                        Args = JsonSerializer.Serialize(
                            new { workbookId = selected.GetProperty("workbookId").GetString() },
                            ServiceProtocol.JsonOptions)
                    });
                    sessionId = GetSessionId(attach);
                }

                sessions[sessionId] = 0;
                await CloseSessionAsync(service, sessionId, save: false);
                sessions.TryRemove(sessionId, out _);
            }
            catch (Exception ex)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, ex);
            }
        }

        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();
    }
}
