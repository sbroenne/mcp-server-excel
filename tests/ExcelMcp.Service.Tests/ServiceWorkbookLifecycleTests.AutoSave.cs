using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Fact]
    public async Task SetAutoSave_ReportsNativeSettingAndLeavesOtherWorkbookUnchanged()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var targetId = await CreateSessionAsync(service, Path.Join(directory, "autosave-target.xlsx"));
            sessions[targetId] = 0;
            var otherId = await CreateSessionAsync(service, Path.Join(directory, "autosave-other.xlsx"));
            sessions[otherId] = 0;
            var target = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(targetId));
            var other = Assert.IsAssignableFrom<IExcelBatch>(service.SessionManager.GetSession(otherId));
            var targetBefore = target.Execute((context, _) =>
                ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn));
            var otherBefore = other.Execute((context, _) =>
                ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn));
            Assert.NotNull(targetBefore);
            Assert.NotNull(otherBefore);
            var requested = !targetBefore.Value;

            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.set-auto-save",
                SessionId = targetId,
                Args = JsonSerializer.Serialize(new { enabled = requested }, ServiceProtocol.JsonOptions)
            });

            var targetAfter = target.Execute((context, _) =>
                ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn));
            if (response.Success)
            {
                using var result = JsonDocument.Parse(response.Result!);
                Assert.Equal(requested, result.RootElement.GetProperty("autoSaveOn").GetBoolean());
                Assert.Equal(requested, targetAfter);
            }
            else
            {
                Assert.NotNull(response.HResult);
                Assert.Equal(targetBefore, targetAfter);
            }
            Assert.Equal(otherBefore, other.Execute((context, _) =>
                ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn)));
            Assert.Same(target, service.SessionManager.GetSession(targetId));
            Assert.Same(other, service.SessionManager.GetSession(otherId));
        });
    }

    [Fact]
    public async Task SetAutoSave_ReadOnlyWorkbookRejectsChangeAndRetainsSession()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var path = Path.Join(directory, "autosave-read-only.xlsx");
            var sessionId = await CreateSessionAsync(service, path);
            sessions[sessionId] = 0;
            await WriteMarkerAsync(service, sessionId, "Saved baseline");
            await CloseSessionAsync(service, sessionId, save: true);
            sessions.TryRemove(sessionId, out _);
            sessionId = await OpenSessionAsync(service, path);
            sessions[sessionId] = 0;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(
                service.SessionManager.GetSession(sessionId));
            batch.Execute((context, _) =>
                context.Book.ChangeFileAccess(Excel.XlFileAccess.xlReadOnly, Type.Missing, false));
            Assert.True(batch.Execute((context, _) => context.Book.ReadOnly));
            Assert.True(batch.Execute((context, _) => context.Book.Saved));
            var before = batch.Execute((context, _) =>
                ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn));
            Assert.NotNull(before);

            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.set-auto-save",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { enabled = !before.Value }, ServiceProtocol.JsonOptions)
            });

            Assert.False(response.Success);
            Assert.Contains("read-only", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(sessionId, response.SessionId);
            Assert.Same(batch, service.SessionManager.GetSession(sessionId));
            Assert.Equal("Saved baseline", await ReadMarkerAsync(service, sessionId));
            Assert.True(batch.Execute((context, _) => context.Book.Saved));
            await CloseSessionAsync(service, sessionId, save: false);
            sessions.TryRemove(sessionId, out _);
            var reopened = await OpenSessionAsync(service, path);
            sessions[reopened] = 0;
            Assert.Equal("Saved baseline", await ReadMarkerAsync(service, reopened));
        });
    }

    [Fact]
    public async Task SetAutoSave_NativeRejectionPreservesUnsavedContentsAndSession()
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            var sessionId = await CreateSessionAsync(service, Path.Join(directory, "autosave-native-rejection.xlsx"));
            sessions[sessionId] = 0;
            var batch = Assert.IsAssignableFrom<IExcelBatch>(
                service.SessionManager.GetSession(sessionId));
            await WriteMarkerAsync(service, sessionId, "Unsaved edit");
            var before = batch.Execute((context, _) =>
                ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn));
            Assert.NotNull(before);

            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.set-auto-save",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { enabled = !before.Value }, ServiceProtocol.JsonOptions)
            });

            Assert.False(response.Success);
            Assert.NotNull(response.HResult);
            Assert.Equal(sessionId, response.SessionId);
            Assert.Same(batch, service.SessionManager.GetSession(sessionId));
            Assert.Equal("Unsaved edit", await ReadMarkerAsync(service, sessionId));
            Assert.False(batch.Execute((context, _) => context.Book.Saved));
            Assert.Equal(before, batch.Execute((context, _) =>
                ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn)));
        });
    }
}
